const XLSX = require("xlsx");

function normaliseCategoryText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function loadCategoryRules(workbookPath, validCategoryIds = null) {
  const workbook = XLSX.readFile(workbookPath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  const rules = [];

  for (const row of rows) {
    const categoryId = String(row.category_id || "").trim();
    const categoryPath = String(row["Category Path"] || "").trim();
    const keywords = String(row.Keywords || "")
      .split(";")
      .map(keyword => keyword.trim())
      .filter(Boolean);

    if (!categoryId || !categoryPath || keywords.length === 0) {
      continue;
    }

    if (validCategoryIds && !validCategoryIds.has(categoryId)) {
      console.warn(
        `Ignoring category rule ${categoryId} because it is not in Trade Me Categories.xlsx`
      );
      continue;
    }

    for (const keyword of keywords) {
      const normalisedKeyword = normaliseCategoryText(keyword);

      if (!normalisedKeyword) {
        continue;
      }

      rules.push({
        categoryId,
        categoryPath,
        keyword,
        normalisedKeyword,
        wordCount: normalisedKeyword.split(" ").length
      });
    }
  }

  // Specific phrases must beat general phrases, for example
  // "xbox series s console" before "xbox console".
  return rules.sort((a, b) =>
    b.wordCount - a.wordCount
    || b.normalisedKeyword.length - a.normalisedKeyword.length
  );
}

function findCategoryMatch(productText, rules) {
  const normalisedText = normaliseCategoryText(productText);

  if (!normalisedText) {
    return null;
  }

  const searchableText = ` ${normalisedText} `;

  for (const rule of rules) {
    if (searchableText.includes(` ${rule.normalisedKeyword} `)) {
      return {
        categoryId: rule.categoryId,
        categoryPath: rule.categoryPath,
        keyword: rule.keyword
      };
    }
  }

  return null;
}

module.exports = {
  findCategoryMatch,
  loadCategoryRules,
  normaliseCategoryText
};
