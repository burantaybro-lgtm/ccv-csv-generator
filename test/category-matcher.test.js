const assert = require("node:assert/strict");
const test = require("node:test");

const {
  findCategoryMatch,
  normaliseCategoryText
} = require("../category-matcher");

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
