const assert = require("node:assert/strict");
const test = require("node:test");

const {
  findCategoryMatch,
  findDepartmentMatch,
  getDepartmentCategoryText,
  loadCategoryRuleWorkbook,
  normaliseCategoryText
} = require("../category-matcher");
const XLSX = require("xlsx");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const rules = [
  {
    categoryId: "5470",
    categoryPath: "Gaming > Xbox-Series-XS > Consoles",
    keyword: "xbox series s console",
    normalisedKeyword: "xbox series s console",
    wordCount: 4
  },
  {
    categoryId: "218",
    categoryPath: "Gaming > Xbox > Consoles",
    keyword: "xbox console",
    normalisedKeyword: "xbox console",
    wordCount: 2
  },
  {
    categoryId: "7838",
    categoryPath: "Gaming > Nintendo-DS > Consoles",
    keyword: "DS console",
    normalisedKeyword: "ds console",
    wordCount: 2
  }
];

test("normalises punctuation and spacing for category matching", () => {
  assert.equal(
    normaliseCategoryText("Microsoft  Xbox-Series S: Console"),
    "microsoft xbox series s console"
  );
});

test("uses the most specific approved keyword", () => {
  const match = findCategoryMatch("MICROSOFT XBOX SERIES S CONSOLE", rules);

  assert.equal(match.categoryId, "5470");
  assert.equal(match.keyword, "xbox series s console");
});

test("does not match DS console inside 3DS console", () => {
  const match = findCategoryMatch("NINTENDO 3DS CONSOLE", rules);
  assert.equal(match, null);
});

test("returns null when no approved keyword matches", () => {
  assert.equal(findCategoryMatch("Makita cordless drill", rules), null);
});

test("loads only approved multi-tab rules and routes by department", () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Department Routing"],
      [],
      ["Department", "Routing Keywords"],
      ["Gaming", "console; playstation"],
      ["Tools & Hardware", "cordless drill; power tool"]
    ]),
    "Department Routing"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Gaming"],
      [],
      [],
      ["category_id", "Category Path", "Category Name", "Keywords", "Status"],
      ["9908", "Gaming > PlayStation-4 > Consoles", "Consoles", "ps4 console", "Approved"],
      ["9909", "Gaming > PlayStation-4 > Games", "Games", "ps4 game", "Add keywords"]
    ]),
    "Gaming"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Tools"],
      [],
      [],
      ["category_id", "Category Path", "Category Name", "Keywords", "Status"],
      ["6020", "Building > Tools > Saws", "Saws", "circular saw", "Approved"]
    ]),
    "Tools & Hardware"
  );

  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), "ccv-rules-"));
  const workbookPath = path.join(temporaryDirectory, "rules.xlsx");
  XLSX.writeFile(workbook, workbookPath);

  const loaded = loadCategoryRuleWorkbook(
    workbookPath,
    new Set(["9908", "9909", "6020"])
  );

  assert.equal(loaded.allRules.length, 2);
  assert.equal(loaded.departments.Gaming.length, 1);
  assert.equal(
    findDepartmentMatch("Sony PlayStation console", loaded).department,
    "Gaming"
  );
  assert.match(
    getDepartmentCategoryText("Gaming", loaded),
    /Category Code: 9908/
  );
  assert.doesNotMatch(
    getDepartmentCategoryText("Gaming", loaded),
    /Category Code: 9909/
  );
});

test("drops an ambiguous approved keyword within one department", () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Department", "Routing Keywords"],
      ["Gaming", "console"]
    ]),
    "Department Routing"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["category_id", "Category Path", "Keywords", "Status"],
      ["1", "Gaming > Gameboy", "gamecube game", "Approved"],
      ["2", "Gaming > GameCube", "gamecube game", "Approved"]
    ]),
    "Gaming"
  );

  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), "ccv-rules-"));
  const workbookPath = path.join(temporaryDirectory, "rules.xlsx");
  XLSX.writeFile(workbook, workbookPath);
  const loaded = loadCategoryRuleWorkbook(workbookPath, new Set(["1", "2"]));

  assert.equal(loaded.departments.Gaming.length, 0);
});
