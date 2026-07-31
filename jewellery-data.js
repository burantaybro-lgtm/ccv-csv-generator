const XLSX = require("xlsx");

const METAL_CODES = [
  ["ST/ST", "Stainless Steel"],
  ["TUNG", "Tungsten"],
  ["SS", "Sterling Silver"],
  ["PT", "Platinum"],
  ["YG", "Yellow Gold"],
  ["WG", "White Gold"],
  ["RG", "Rose Gold"],
  ["TI", "Titanium"]
];

function normaliseJewelleryStockCode(value) {
  const match = String(value || "").trim().match(/\b[AB]\d+-\d+\b/i);
  return match ? match[0].toUpperCase() : null;
}

function extractMetal(description) {
  const text = String(description || "").toUpperCase();

  for (const [code, name] of METAL_CODES) {
    const escapedCode = code.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const match = text.match(
      new RegExp(`\\b(\\d{1,2}CT\\s+)?${escapedCode}\\b`)
    );

    if (match) {
      return `${match[1] || ""}${name}`.trim();
    }
  }

  return "";
}

function extractWeight(description) {
  const match = String(description || "").match(
    /\bTW\s*([\d.]+)\s*(?:GMS?|GRAMS?)\b/i
  );

  return match ? `${match[1]} grams` : "";
}

function parseJewelleryReport(fileBuffer, reportFilename) {
  const workbook = XLSX.read(fileBuffer, {
    type: "buffer",
    cellDates: false
  });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: true
  });
  const headerIndex = rows.findIndex(row =>
    row.some(cell => /^Stock Code$/i.test(String(cell || "").trim()))
    && row.some(cell => /^Description$/i.test(String(cell || "").trim()))
  );

  if (headerIndex < 0) {
    return {
      products: [],
      warnings: [`${reportFilename}: Stock Code and Description columns were not found`]
    };
  }

  const headers = rows[headerIndex].map(value => String(value || "").trim());
  const stockCodeColumn = headers.findIndex(header => /^Stock Code$/i.test(header));
  const descriptionColumn = headers.findIndex(header => /^Description$/i.test(header));
  const products = [];
  let currentProduct = null;

  for (const row of rows.slice(headerIndex + 1)) {
    const stockCode = normaliseJewelleryStockCode(row[stockCodeColumn]);
    const description = String(row[descriptionColumn] || "").trim();

    if (stockCode) {
      currentProduct = {
        stockCode,
        jewelleryDescription: description,
        sourceJewelleryReport: reportFilename
      };
      products.push(currentProduct);
      continue;
    }

    if (
      currentProduct
      && description
      && !/^Item Count:/i.test(description)
      && !/^Total Cost:/i.test(description)
    ) {
      currentProduct.jewelleryDescription =
        `${currentProduct.jewelleryDescription} ${description}`.trim();
    }
  }

  for (const product of products) {
    product.isWatch = /\bWATCH\b/i.test(product.jewelleryDescription);
    product.metal = extractMetal(product.jewelleryDescription);
    product.weight = extractWeight(product.jewelleryDescription);
  }

  return {
    products,
    warnings: products.length
      ? []
      : [`${reportFilename}: no jewellery stock items were found`]
  };
}

function mergeJewelleryProducts(target, products) {
  const result = target || {};

  for (const product of products) {
    result[product.stockCode] = {
      ...result[product.stockCode],
      ...product
    };
  }

  return result;
}

function mergeJewelleryScreenDetails(target, details) {
  const result = target || {};

  for (const detail of details || []) {
    const stockCode = normaliseJewelleryStockCode(detail.stockCode);

    if (!stockCode) {
      continue;
    }

    result[stockCode] = {
      ...result[stockCode],
      stockCode,
      size: String(detail.size || "").replace(/^SIZE\s*/i, "").trim(),
      ccCode: String(detail.ccCode || "").trim(),
      screenConfidence: String(detail.confidence || "").toLowerCase()
    };
  }

  return result;
}

function getJewelleryReviewIssue(record) {
  if (!record || record.isWatch) {
    return null;
  }

  if (!record.ccCode) {
    return "CC Code could not be read from the jewellery details screenshot";
  }

  if (
    /\bRING\b/i.test(record.jewelleryDescription || "")
    && !record.size
  ) {
    return "Ring size could not be read from the jewellery details screenshot";
  }

  if (record.screenConfidence && record.screenConfidence !== "high") {
    return "Jewellery screenshot details were not read with high confidence";
  }

  return null;
}

function getValuedTitleParts(listing) {
  const valueText = String(listing?.new_retail_price || "").trim();
  const numericValue = Number(valueText.replace(/[$,\s]/g, ""));
  const newRetailValue = Number.isFinite(numericValue) && numericValue > 0
    ? `$${numericValue.toLocaleString("en-NZ", { maximumFractionDigits: 2 })}`
    : "";
  const metalText = [
    listing?.metal,
    listing?.hallmark
  ].filter(Boolean).join(" ").toUpperCase();
  const caratMatch = metalText.match(/\b(\d{1,2})\s*(?:CT|CARAT|K)\b/);
  const hallmark = String(listing?.hallmark || "").replace(/\D/g, "");
  const hallmarkCarats = {
    375: "9CT",
    417: "10CT",
    585: "14CT",
    750: "18CT",
    916: "22CT",
    917: "22CT"
  };
  const carat = caratMatch
    ? `${Number(caratMatch[1])}CT`
    : hallmarkCarats[hallmark] || "";
  let colour = "";
  let metalLabel = "";
  const isGold = /\bGOLD\b|\bYG\b|\bWG\b|\bRG\b/.test(metalText)
    || Boolean(hallmarkCarats[hallmark]);

  if (/\bYG\b|YELLOW GOLD/.test(metalText)) {
    colour = "YG";
  } else if (/\bWG\b|WHITE GOLD/.test(metalText)) {
    colour = "WG";
  } else if (/\bRG\b|ROSE GOLD/.test(metalText)) {
    colour = "RG";
  }

  if (!isGold) {
    if (/\bPLATINUM\b|\bPT\b/.test(metalText)) {
      metalLabel = "Platinum";
    } else if (/\bPALLADIUM\b|\bPD\b/.test(metalText)) {
      metalLabel = "Palladium";
    } else if (/\bSTERLING SILVER\b|\bSS\b/.test(metalText) || hallmark === "925") {
      metalLabel = "Sterling Silver";
    } else if (/\bFINE SILVER\b/.test(metalText) || hallmark === "999") {
      metalLabel = "Fine Silver";
    } else if (/\bSILVER\b/.test(metalText)) {
      metalLabel = "Silver";
    } else if (/\bTITANIUM\b|\bTI\b/.test(metalText)) {
      metalLabel = "Titanium";
    } else if (/\bSTAINLESS STEEL\b|\bST\/ST\b/.test(metalText)) {
      metalLabel = "Stainless Steel";
    } else if (/\bTUNGSTEN\b|\bTUNG\b/.test(metalText)) {
      metalLabel = "Tungsten";
    }
  }

  return {
    newRetailValue,
    isGold,
    carat: isGold ? carat : "",
    colour: isGold ? colour : "",
    metalLabel
  };
}

function buildValuedJewelleryTitle(listing) {
  const {
    newRetailValue,
    isGold,
    carat,
    colour,
    metalLabel
  } = getValuedTitleParts(listing);
  const cleanTitle = String(listing?.title || "")
    .replace(/^\s*\$[\d,.]+\s*/i, "")
    .replace(/^\s*\d{1,2}\s*(?:CT|CARAT|K)\s+(?:YG|WG|RG)\b\s*/i, "")
    .replace(
      /^\s*(?:Platinum|Palladium|Sterling Silver|Fine Silver|Silver|Titanium|Stainless Steel|Tungsten)\b\s*/i,
      ""
    )
    .trim();
  const titleMetal = isGold
    ? [carat, colour].filter(Boolean).join(" ")
    : metalLabel;

  return [newRetailValue, titleMetal, cleanTitle]
    .filter(Boolean)
    .join(" ");
}

module.exports = {
  buildValuedJewelleryTitle,
  extractMetal,
  extractWeight,
  getJewelleryReviewIssue,
  getValuedTitleParts,
  mergeJewelleryProducts,
  mergeJewelleryScreenDetails,
  normaliseJewelleryStockCode,
  parseJewelleryReport
};
