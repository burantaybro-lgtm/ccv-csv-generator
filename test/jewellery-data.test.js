const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const XLSX = require("xlsx");

const {
  buildValuedJewelleryTitle,
  extractMetal,
  extractWeight,
  getJewelleryReviewIssue,
  getValuedTitleParts,
  mergeJewelleryProducts,
  mergeJewelleryScreenDetails,
  parseJewelleryReport
} = require("../jewellery-data");

test("expands jewellery metal codes and extracts total weight", () => {
  const description =
    "9CT YG DRESS RING DIAMOND CLUSTER 1X2PT ROUND DIAMOND TW 2.15GMS";

  assert.equal(extractMetal(description), "9CT Yellow Gold");
  assert.equal(extractWeight(description), "2.15 grams");
  assert.equal(extractMetal("SS CHAIN TW 15GMS"), "Sterling Silver");
});

test("parses wrapped jewellery descriptions from the supplied legacy XLS", () => {
  const reportPath = path.join(
    __dirname,
    "..",
    "..",
    "..",
    "upload",
    "UNVALUED JEWELLERY(1).xls"
  );

  if (!fs.existsSync(reportPath)) {
    return;
  }

  const parsed = parseJewelleryReport(
    fs.readFileSync(reportPath),
    "UNVALUED JEWELLERY.xls"
  );
  const product = parsed.products.find(
    item => item.stockCode === "A18109224-1"
  );

  assert.equal(parsed.products.length, 17);
  assert.match(product.jewelleryDescription, /45X0\.5PT ROUND DIAMOND/);
  assert.equal(product.metal, "9CT Yellow Gold");
  assert.equal(product.weight, "2.15 grams");
});

test("parses a synthetic report and joins continuation rows", () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Jewellery Awaiting Processing"],
      [],
      ["Stock Code", "", "Description", "Cost"],
      ["B18194037-1", "", "9CT YG RING 1X50PT OVAL GREEN STONE", 200],
      ["", "", "6X3PT ROUND DIAMOND TW 5.9GMS", ""]
    ]),
    "Sheet1"
  );
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "biff8" });
  const parsed = parseJewelleryReport(buffer, "test.xls");

  assert.equal(parsed.products.length, 1);
  assert.equal(parsed.products[0].metal, "9CT Yellow Gold");
  assert.equal(parsed.products[0].weight, "5.9 grams");
  assert.match(parsed.products[0].jewelleryDescription, /6X3PT ROUND DIAMOND/);
});

test("merges screenshot size and CC Code by stock code", () => {
  let products = mergeJewelleryProducts({}, [{
    stockCode: "A18109224-1",
    jewelleryDescription: "9CT YG DRESS RING TW 2.15GMS",
    isWatch: false
  }]);

  products = mergeJewelleryScreenDetails(products, [{
    stockCode: "A18109224-1",
    size: "SIZE O",
    ccCode: "202833",
    confidence: "high"
  }]);

  assert.equal(products["A18109224-1"].size, "O");
  assert.equal(products["A18109224-1"].ccCode, "202833");
  assert.equal(getJewelleryReviewIssue(products["A18109224-1"]), null);
});

test("requires a readable size for rings and CC Code for standard jewellery", () => {
  assert.match(
    getJewelleryReviewIssue({
      jewelleryDescription: "9CT YG RING TW 2GMS",
      ccCode: "202833",
      size: "",
      screenConfidence: "high"
    }),
    /Ring size/
  );

  assert.match(
    getJewelleryReviewIssue({
      jewelleryDescription: "SS CHAIN TW 15GMS",
      ccCode: "",
      size: "",
      screenConfidence: "high"
    }),
    /CC Code/
  );
});

test("builds a valued jewellery title from retail value, hallmark and colour", () => {
  const listing = {
    title: "Sapphire & Diamond Cluster Ring",
    new_retail_price: "5600",
    hallmark: "Stamped 750",
    metal: "18 carat yellow gold"
  };

  assert.deepEqual(getValuedTitleParts(listing), {
    newRetailValue: "$5,600",
    isGold: true,
    carat: "18CT",
    colour: "YG",
    metalLabel: ""
  });
  assert.equal(
    buildValuedJewelleryTitle(listing),
    "$5,600 18CT YG Sapphire & Diamond Cluster Ring"
  );
});

test("does not duplicate an existing valued jewellery title prefix", () => {
  const title = buildValuedJewelleryTitle({
    title: "$5,000 18CT YG Diamond Dress Ring",
    new_retail_price: "$5,000",
    hallmark: "750",
    metal: "Yellow Gold"
  });

  assert.equal(title, "$5,000 18CT YG Diamond Dress Ring");
});

test("uses the metal name instead of gold fields for valued platinum jewellery", () => {
  const listing = {
    title: "Diamond Solitaire Ring",
    new_retail_price: "5000",
    hallmark: "950",
    metal: "Platinum"
  };

  assert.deepEqual(getValuedTitleParts(listing), {
    newRetailValue: "$5,000",
    isGold: false,
    carat: "",
    colour: "",
    metalLabel: "Platinum"
  });
  assert.equal(
    buildValuedJewelleryTitle(listing),
    "$5,000 Platinum Diamond Solitaire Ring"
  );
});

test("supports other non-gold metals in valued jewellery titles", () => {
  assert.equal(
    buildValuedJewelleryTitle({
      title: "Sterling Silver Dress Ring",
      new_retail_price: "1800",
      hallmark: "925",
      metal: "Sterling Silver"
    }),
    "$1,800 Sterling Silver Dress Ring"
  );
});
