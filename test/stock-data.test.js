const assert = require("node:assert/strict");
const test = require("node:test");
const XLSX = require("xlsx");

const {
  createEmptyProductDatabase,
  getStockCodeFromFilename,
  mergeProducts,
  parseBuyReport,
  parseStockReport
} = require("../stock-data");

test("groups numbered photo filenames under one stock code", () => {
  assert.equal(getStockCodeFromFilename("B18194032-2.jpg"), "B18194032-2");
  assert.equal(getStockCodeFromFilename("B18194032-2 (1).jpg"), "B18194032-2");
  assert.equal(getStockCodeFromFilename("b18194032-2 (12).PNG"), "B18194032-2");
  assert.equal(getStockCodeFromFilename("A2149497-1 (2).jpg"), "A2149497-1");
  assert.equal(getStockCodeFromFilename("not-a-stock-code.jpg"), null);
});

test("parses A-stock loan reports including continued descriptions", () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ["Pulled Cash Loans"],
    ["From 23/07/2026 to 23/07/2026"],
    [],
    ["CL #", "Date", "Due", "Bay", "Amount"],
    [2149497, new Date("2026-04-06T11:20:48Z"), new Date("2026-07-11T00:00:00Z"), "PRIVATE", 120],
    ["A2149497-1 SONY ADAPTIVE NOISE CANCELLING WIRELESS OVER-EAR HEADPHONES"],
    ["WH1000XM6 HEADPHONES"]
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "biff8" });
  const parsed = parseStockReport(buffer, "2026-07-23.xls");

  assert.equal(parsed.warnings.length, 0);
  assert.equal(parsed.products.length, 1);
  assert.equal(parsed.products[0].stockCode, "A2149497-1");
  assert.equal(parsed.products[0].stockType, "loan");
  assert.equal(parsed.products[0].loanNumber, "2149497");
  assert.equal(parsed.products[0].reportDate, "2026-07-23");
  assert.match(parsed.products[0].originalDescription, /WH1000XM6 HEADPHONES$/);
  assert.equal(JSON.stringify(parsed.products).includes("PRIVATE"), false);
  assert.equal(JSON.stringify(parsed.products).includes("120"), false);
});

test("parses a legacy XLS buy report and warns about filename date", () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ["Buys by Date"],
    ["From 9/07/2026 to 9/07/2026"],
    ["Buy #", "Date", "Amount", "Customer"],
    [18194032, new Date("2026-07-09T11:01:20Z"), 120, "PRIVATE"],
    ["B18194032-1 NIKON COOLPIX L20 DIGITAL CAMERA"],
    ["B18194032-2 CANON EOS 700D DS126431 DIGITAL CAMERA"]
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "biff8" });
  const parsed = parseBuyReport(buffer, "2027-07-09.xls");

  assert.equal(parsed.products.length, 2);
  assert.equal(parsed.products[0].stockCode, "B18194032-1");
  assert.equal(parsed.products[1].stockCode, "B18194032-2");
  assert.equal(parsed.products[1].buyNumber, "18194032");
  assert.match(parsed.warnings.join("\n"), /does not match report date 2026-07-09/);
  assert.equal(JSON.stringify(parsed.products).includes("PRIVATE"), false);
});

test("imports an xlsx report and updates existing database records", () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ["Buys by Date"],
    ["From 23/07/2026 to 23/07/2026"],
    ["Buy #", "Date", "Amount", "Customer"],
    [18199999, new Date("2026-07-23T10:00:00Z"), 50, "PRIVATE"],
    ["B18199999-1 MAKITA DHP484 CORDLESS DRILL"]
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  const parsed = parseBuyReport(buffer, "2026-07-23.xlsx");
  const database = mergeProducts(createEmptyProductDatabase(), parsed.products);

  assert.equal(parsed.warnings.length, 0);
  assert.equal(database.products["B18199999-1"].buyNumber, "18199999");
  assert.equal(
    database.products["B18199999-1"].originalDescription,
    "MAKITA DHP484 CORDLESS DRILL"
  );
  assert.equal(JSON.stringify(database).includes("PRIVATE"), false);
});
