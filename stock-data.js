const path = require("path");
const XLSX = require("xlsx");

function normaliseStockCode(value) {
  const match = String(value || "").trim().match(/\bB\d+-\d+\b/i);
  return match ? match[0].toUpperCase() : null;
}

function getStockCodeFromFilename(filename) {
  const name = path.parse(String(filename || "")).name;
  const withoutPhotoNumber = name.replace(/\s*\(\d+\)\s*$/, "");
  return normaliseStockCode(withoutPhotoNumber);
}

function excelDateToIso(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString();
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);

    if (parsed) {
      return new Date(Date.UTC(
        parsed.y,
        parsed.m - 1,
        parsed.d,
        parsed.H || 0,
        parsed.M || 0,
        Math.floor(parsed.S || 0)
      )).toISOString();
    }
  }

  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed.toISOString();
}

function getDateOnly(isoDate) {
  return isoDate ? isoDate.slice(0, 10) : null;
}

function parseBuyReport(fileBuffer, reportFilename) {
  const workbook = XLSX.read(fileBuffer, {
    type: "buffer",
    // Keep Excel serial dates so the report's calendar date is not shifted by
    // the Render server timezone when it is converted to JavaScript Date.
    cellDates: false
  });

  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: null,
    raw: true
  });

  const products = [];
  const warnings = [];
  let currentBuy = null;

  for (const row of rows) {
    const firstCell = String(row[0] ?? "").trim();

    if (/^\d{7,}$/.test(firstCell) && row[1]) {
      currentBuy = {
        buyNumber: firstCell,
        transactionDate: excelDateToIso(row[1])
      };
      continue;
    }

    const itemMatch = firstCell.match(/^(B\d+-\d+)\s+(.+)$/i);

    if (!itemMatch) {
      continue;
    }

    products.push({
      stockCode: itemMatch[1].toUpperCase(),
      originalDescription: itemMatch[2].trim(),
      buyNumber: currentBuy?.buyNumber || null,
      transactionDate: currentBuy?.transactionDate || null,
      sourceReport: reportFilename
    });
  }

  if (products.length === 0) {
    warnings.push(`${reportFilename}: no stock items were found`);
  }

  const filenameDate = path.parse(reportFilename).name.match(/^\d{4}-\d{2}-\d{2}$/)?.[0];
  const reportDate = getDateOnly(products.find(product => product.transactionDate)?.transactionDate);

  if (filenameDate && reportDate && filenameDate !== reportDate) {
    warnings.push(
      `${reportFilename}: filename date ${filenameDate} does not match report date ${reportDate}`
    );
  }

  return { products, warnings };
}

function createEmptyProductDatabase() {
  return {
    version: 1,
    updatedAt: null,
    products: {}
  };
}

function mergeProducts(database, products) {
  const target = database && typeof database === "object"
    ? database
    : createEmptyProductDatabase();

  target.version = 1;
  target.products = target.products || {};

  for (const product of products) {
    target.products[product.stockCode] = {
      ...target.products[product.stockCode],
      ...product,
      updatedAt: new Date().toISOString()
    };
  }

  target.updatedAt = new Date().toISOString();
  return target;
}

module.exports = {
  createEmptyProductDatabase,
  getStockCodeFromFilename,
  mergeProducts,
  normaliseStockCode,
  parseBuyReport
};
