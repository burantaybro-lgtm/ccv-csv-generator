const path = require("path");
const XLSX = require("xlsx");

function normaliseStockCode(value) {
  const match = String(value || "").trim().match(/\b[AB]\d+-\d+\b/i);
  return match ? match[0].toUpperCase() : null;
}

function getStockCodeFromFilename(filename) {
  const name = path.parse(String(filename || "")).name;
  const withoutPhotoNumber = name.replace(/\s*\(\d+\)\s*$/, "");
  return normaliseStockCode(withoutPhotoNumber);
}

function buildTradeMeTitle(title, stockCode, maximumLength = 80) {
  const cleanStockCode = String(stockCode || "").trim().toUpperCase();
  const stockTag = cleanStockCode ? `#${cleanStockCode}` : "";
  const escapedTag = stockTag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let baseTitle = String(title || "")
    .replace(new RegExp(`\\s*${escapedTag}`, "ig"), "")
    .replace(/\s+/g, " ")
    .trim();

  if (!stockTag) {
    return baseTitle.slice(0, maximumLength).trim();
  }

  const availableTitleLength = maximumLength - stockTag.length - 1;

  if (availableTitleLength <= 0) {
    return stockTag.slice(0, maximumLength);
  }

  if (baseTitle.length > availableTitleLength) {
    const hardCut = baseTitle.slice(0, availableTitleLength).trim();
    const lastSpace = hardCut.lastIndexOf(" ");

    baseTitle = lastSpace >= Math.floor(availableTitleLength * 0.6)
      ? hardCut.slice(0, lastSpace)
      : hardCut;
  }

  return `${baseTitle} ${stockTag}`.trim();
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
      stockType: "buy",
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

function parseNzReportDate(rows) {
  for (const row of rows.slice(0, 5)) {
    const text = String(row[0] || "");
    const match = text.match(/From\s+(\d{1,2})\/(\d{1,2})\/(\d{4})/i);

    if (match) {
      const [, day, month, year] = match;
      return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    }
  }

  return null;
}

function parseLoanReport(fileBuffer, reportFilename) {
  const workbook = XLSX.read(fileBuffer, {
    type: "buffer",
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
  const reportDate = parseNzReportDate(rows);
  let currentLoan = null;
  let currentProduct = null;

  for (const row of rows) {
    const firstCell = String(row[0] ?? "").trim();

    if (/^\d{7,}$/.test(firstCell) && row[1]) {
      currentLoan = {
        loanNumber: firstCell,
        transactionDate: excelDateToIso(row[1]),
        dueDate: excelDateToIso(row[2])
      };
      currentProduct = null;
      continue;
    }

    const itemMatch = firstCell.match(/^(A\d+-\d+)\s+(.+)$/i);

    if (itemMatch) {
      currentProduct = {
        stockCode: itemMatch[1].toUpperCase(),
        stockType: "loan",
        originalDescription: itemMatch[2].trim(),
        loanNumber: currentLoan?.loanNumber || null,
        transactionDate: currentLoan?.transactionDate || null,
        dueDate: currentLoan?.dueDate || null,
        reportDate,
        sourceReport: reportFilename
      };
      products.push(currentProduct);
      continue;
    }

    const isContinuation = currentProduct
      && firstCell
      && row.slice(1).every(value => value === null)
      && !/^Pulled Cash Loans$/i.test(firstCell)
      && !/^From\s+/i.test(firstCell)
      && !/^CL\s*#/i.test(firstCell)
      && !/^\d{1,2}\/\d{1,2}\/\d{4}/.test(firstCell);

    if (isContinuation) {
      currentProduct.originalDescription += ` ${firstCell}`;
    }
  }

  if (products.length === 0) {
    warnings.push(`${reportFilename}: no A-stock loan items were found`);
  }

  const filenameDate = path.parse(reportFilename).name.match(/^\d{4}-\d{2}-\d{2}$/)?.[0];

  if (filenameDate && reportDate && filenameDate !== reportDate) {
    warnings.push(
      `${reportFilename}: filename date ${filenameDate} does not match report date ${reportDate}`
    );
  }

  return { products, warnings };
}

function parseStockReport(fileBuffer, reportFilename) {
  const workbook = XLSX.read(fileBuffer, {
    type: "buffer",
    cellDates: false,
    sheetRows: 5
  });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: null,
    raw: true
  });
  const heading = rows.flat().filter(Boolean).join(" ");

  if (/Pulled Cash Loans/i.test(heading)) {
    return parseLoanReport(fileBuffer, reportFilename);
  }

  if (/Buys by Date/i.test(heading)) {
    return parseBuyReport(fileBuffer, reportFilename);
  }

  return {
    products: [],
    warnings: [`${reportFilename}: unsupported report layout`]
  };
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
  buildTradeMeTitle,
  createEmptyProductDatabase,
  getStockCodeFromFilename,
  mergeProducts,
  normaliseStockCode,
  parseBuyReport,
  parseLoanReport,
  parseStockReport
};
