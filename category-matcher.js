const XLSX = require("xlsx");

const IGNORED_SHEETS = new Set(["Instructions", "Department Routing"]);

function normaliseCategoryText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sortRules(rules) {
  return rules.sort((a, b) =>
    b.wordCount - a.wordCount
    || b.normalisedKeyword.length - a.normalisedKeyword.length
  );
}

function findHeaderRow(rows, requiredColumn) {
  return rows.findIndex(row =>
    row.some(cell => String(cell || "").trim() === requiredColumn)
  );
}

function rowsToObjects(rows, headerIndex) {
  const headers = rows[headerIndex].map(value => String(value || "").trim());

  return rows.slice(headerIndex + 1).map(row =>
    Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""]))
  );
}

function loadCategoryRuleWorkbook(workbookPath, validCategoryIds = null) {
  const workbook = XLSX.readFile(workbookPath);
  const departments = {};
  const allRules = [];

  for (const sheetName of workbook.SheetNames) {
    if (IGNORED_SHEETS.has(sheetName)) {
      continue;
    }

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: ""
    });
    const headerIndex = findHeaderRow(rows, "category_id");

    if (headerIndex < 0) {
      continue;
    }

    const rules = [];

    for (const row of rowsToObjects(rows, headerIndex)) {
      const categoryId = String(row.category_id || "").trim();
      const categoryPath = String(row["Category Path"] || "").trim();
      const status = String(row.Status || "").trim().toLowerCase();
      const keywords = String(row.Keywords || "")
        .split(";")
        .map(keyword => keyword.trim())
        .filter(Boolean);

      if (
        status !== "approved"
        || !categoryId
        || !categoryPath
        || keywords.length === 0
      ) {
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
          department: sheetName,
          categoryId,
          categoryPath,
          keyword,
          normalisedKeyword,
          wordCount: normalisedKeyword.split(" ").length
        });
      }
    }

    // If an approved keyword points to two categories in the same department,
    // exclude it instead of making an arbitrary choice.
    const categoryIdsByKeyword = new Map();
    for (const rule of rules) {
      if (!categoryIdsByKeyword.has(rule.normalisedKeyword)) {
        categoryIdsByKeyword.set(rule.normalisedKeyword, new Set());
      }
      categoryIdsByKeyword.get(rule.normalisedKeyword).add(rule.categoryId);
    }

    const safeRules = rules.filter(rule =>
      categoryIdsByKeyword.get(rule.normalisedKeyword).size === 1
    );

    departments[sheetName] = sortRules(safeRules);
    allRules.push(...safeRules);
  }

  const routingSheet = workbook.Sheets["Department Routing"];
  const routingRules = [];

  if (routingSheet) {
    const rows = XLSX.utils.sheet_to_json(routingSheet, {
      header: 1,
      defval: ""
    });
    const headerIndex = findHeaderRow(rows, "Department");

    if (headerIndex >= 0) {
      for (const row of rowsToObjects(rows, headerIndex)) {
        const department = String(row.Department || "").trim();

        if (!department || !departments[department]) {
          continue;
        }

        const keywords = String(row["Routing Keywords"] || "")
          .split(";")
          .map(keyword => keyword.trim())
          .filter(Boolean);

        for (const keyword of keywords) {
          const normalisedKeyword = normaliseCategoryText(keyword);
          if (normalisedKeyword) {
            routingRules.push({
              department,
              keyword,
              normalisedKeyword,
              wordCount: normalisedKeyword.split(" ").length
            });
          }
        }
      }
    }
  }

  return {
    allRules: sortRules(allRules),
    departments,
    routingRules: sortRules(routingRules)
  };
}

// Backwards-compatible loader for the original one-sheet rules workbook.
function loadCategoryRules(workbookPath, validCategoryIds = null) {
  return loadCategoryRuleWorkbook(workbookPath, validCategoryIds).allRules;
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
        department: rule.department,
        categoryId: rule.categoryId,
        categoryPath: rule.categoryPath,
        keyword: rule.keyword
      };
    }
  }

  return null;
}

function findDepartmentMatch(productText, categoryRuleWorkbook) {
  const routingMatch = findCategoryMatch(
    productText,
    categoryRuleWorkbook.routingRules.map(rule => ({
      ...rule,
      categoryId: "",
      categoryPath: ""
    }))
  );

  if (routingMatch) {
    return {
      department: routingMatch.department,
      keyword: routingMatch.keyword,
      source: "routing keyword"
    };
  }
  return null;
}

function getDepartmentCategoryText(department, categoryRuleWorkbook) {
  const rules = categoryRuleWorkbook.departments[department] || [];
  const categories = new Map();

  for (const rule of rules) {
    categories.set(
      rule.categoryId,
      `Category Code: ${rule.categoryId} | Path: ${rule.categoryPath}`
    );
  }

  return [...categories.values()].join("\n");
}

module.exports = {
  findCategoryMatch,
  findDepartmentMatch,
  getDepartmentCategoryText,
  loadCategoryRules,
  loadCategoryRuleWorkbook,
  normaliseCategoryText
};
