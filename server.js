require("dotenv").config();

// CSV generator only: reads photos, creates AI listings, and writes Trade Me CSV files.

const XLSX = require("xlsx");
const express = require("express");
const multer = require("multer");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const OpenAI = require("openai");
const { Dropbox } = require("dropbox");
const createCsvWriter = require("csv-writer").createObjectCsvWriter;

const app = express();

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});


async function getDropboxAccessToken() {
  const response = await fetch("https://api.dropboxapi.com/oauth2/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: process.env.DROPBOX_REFRESH_TOKEN,
      client_id: process.env.DROPBOX_CLIENT_ID,
      client_secret: process.env.DROPBOX_CLIENT_SECRET
    })
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error("Dropbox token refresh failed: " + JSON.stringify(data));
  }

  return data.access_token;
}

if (!fs.existsSync("uploads")) {
  fs.mkdirSync("uploads");
}

// LOAD TRADE ME CATEGORY FILE
const workbook = XLSX.readFile("Trade Me Categories.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const categories = XLSX.utils.sheet_to_json(sheet);

app.use(cors());
app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

const storage = multer.diskStorage({
  destination: "uploads/",
  filename: (req, file, cb) => {
    cb(null, Date.now() + path.extname(file.originalname));
  }
});

const upload = multer({ storage });

const csvWriter = createCsvWriter({
  path: "trade-me-auto-listings.csv",

  header: [
    { id: "product_id_for_member", title: "product_id_for_member" },
    { id: "sku", title: "sku" },
    { id: "stock_amount", title: "stock_amount" },
    { id: "unlimited_stock", title: "unlimited_stock" },
    { id: "category_id", title: "category_id" },
    { id: "second_category_id", title: "second_category_id" },
    { id: "title", title: "title" },
    { id: "subtitle", title: "subtitle" },
    { id: "body", title: "body" },
    { id: "is_new", title: "is_new" },
    { id: "attributes", title: "attributes" },
    { id: "is_legal_notice_read", title: "is_legal_notice_read" },
    { id: "start_price", title: "start_price" },
    { id: "reserve_price", title: "reserve_price" },
    { id: "buy_now_price", title: "buy_now_price" },
    { id: "is_clearance", title: "is_clearance" },
    { id: "was_price", title: "was_price" },
    { id: "has_promo", title: "has_promo" },
    { id: "is_sold_in_multiple_quantities", title: "is_sold_in_multiple_quantities" },
    { id: "is_shipping_price_per_quantity_sold", title: "is_shipping_price_per_quantity_sold" },
    { id: "fpo_amount", title: "fpo_amount" },
    { id: "fpo_duration", title: "fpo_duration" },
    { id: "fpo_to", title: "fpo_to" },
    { id: "av_bidders_only", title: "av_bidders_only" },
    { id: "auction_length", title: "auction_length" },
    { id: "auction_end_time", title: "auction_end_time" },
    { id: "delivery_pickup_allowed", title: "delivery_pickup_allowed" },
    { id: "delivery_must_pickup", title: "delivery_must_pickup" },
    { id: "delivery_use_bookcourier_rates", title: "delivery_use_bookcourier_rates" },
    { id: "delivery_bookcourier_is_box", title: "delivery_bookcourier_is_box" },
    { id: "delivery_bookcourier_bag_size", title: "delivery_bookcourier_bag_size" },
    { id: "delivery_bookcourier_selected_courier", title: "delivery_bookcourier_selected_courier" },
    { id: "delivery_bookcourier_service_level", title: "delivery_bookcourier_service_level" },
    { id: "delivery_bookcourier_no_restricted_items", title: "delivery_bookcourier_no_restricted_items" },
    { id: "delivery_price", title: "delivery_price" },
    { id: "payment_bank_deposit", title: "payment_bank_deposit" },
    { id: "payment_credit_card", title: "payment_credit_card" },
    { id: "payment_cash", title: "payment_cash" },
    { id: "payment_afterpay", title: "payment_afterpay" },
    { id: "payment_other", title: "payment_other" },
    { id: "send_payment_instructions", title: "send_payment_instructions" },
    { id: "photo_id_list", title: "photo_id_list" },
    { id: "youtube_video_key", title: "youtube_video_key" },
    { id: "display_bold", title: "display_bold" },
    { id: "gallery", title: "gallery" },
    { id: "gallery_plus", title: "gallery_plus" },
    { id: "feature", title: "feature" },
    { id: "super_feature", title: "super_feature" },
    { id: "donation_recipient", title: "donation_recipient" },
    { id: "folder", title: "folder" },
    { id: "exclude_shipping_promotion", title: "exclude_shipping_promotion" },
    { id: "listing_footer_enabled", title: "listing_footer_enabled" },
    { id: "height_cm", title: "height_cm" },
    { id: "width_cm", title: "width_cm" },
    { id: "length_cm", title: "length_cm" },
    { id: "weight_kg", title: "weight_kg" },
    { id: "brand", title: "brand" },
    { id: "manufacturer_code", title: "manufacturer_code" },
    { id: "barcode_gtin", title: "barcode_gtin" },
    { id: "update_active_listings", title: "update_active_listings" }
  ],

  append: fs.existsSync("trade-me-auto-listings.csv")
});

const searchWords = [
  "watch", "casio", "g-shock", "bike", "phone", "laptop",
  "tool", "drill", "game", "console", "speaker", "headphones",
  "jewellery", "ring", "necklace", "pendant", "earrings",
  "chain", "bracelet", "bangle"
];

const categoryOverrides = [
  {
    keywords: ["mens watch", "men's watch", "dress watch"],
    category_id: "7354"
  },
  {
    keywords: ["yg ring", "yellow gold ring", "diamond ring"],
    category_id: "7374"
  },
  {
    keywords: ["necklace", "chain", "gold chain"],
    category_id: "3999"
  },
  {
    keywords: ["ps4 console", "playstation 4"],
    category_id: "9908"
  },
  {
    keywords: ["ps5 console", "playstation 5"],
    category_id: "5466"
  },
  {
    keywords: ["ps3 console", "playstation 3"],
    category_id: "6207"
  },
  {
    keywords: ["ps2 console", "playstation 2"],
    category_id: "858"
  },
  {
    keywords: ["sewing", "overlocker"],
    category_id: "1099"
  },
  {
    keywords: ["jigsaw", "saw", "circular saw"],
    category_id: "6020"
  }
];

function getCategoryText() {
  return categories
    .filter(cat => {
      const pathText = String(cat["Category Path"] || "").toLowerCase();
      const nameText = String(cat["Name"] || "").toLowerCase();

      return searchWords.some(word =>
        pathText.includes(word) || nameText.includes(word)
      );
    })
    .slice(0, 150)
    .map(cat =>
      `Category Code: ${cat["Category Code"]} | Path: ${cat["Category Path"]}`
    )
    .join("\n");
}

function getDeliveryPrice(listing) {
  const text =
    `${listing.title} ${listing.body}`.toLowerCase();

  if (
    text.includes("watch") ||
    text.includes("ring") ||
    text.includes("necklace")
  ) {
    return "Small bag";
  }

  if (
    text.includes("console") ||
    text.includes("drill") ||
    text.includes("speaker")
  ) {
    return "Shipping box";
  }

  if (
    text.includes("bike") ||
text.includes("guitar") ||
    text.includes("tv")
  ) {
    return "";
  }

  return "";
}

function applyTradeMeDefaults(listing) {
  listing.sku = listing.sku || "";
  listing.stock_amount = "1";
  listing.unlimited_stock = "FALSE";
  listing.second_category_id = "";
  listing.subtitle = "";
  listing.is_new = listing.is_new || "FALSE";
  listing.attributes = "";

  listing.is_legal_notice_read = "TRUE";
  listing.start_price = listing.start_price || "";
  listing.reserve_price = listing.reserve_price || "";
  listing.buy_now_price = listing.buy_now_price || "";
  listing.is_clearance = "FALSE";
  listing.was_price = "";
  listing.has_promo = "N";

  listing.is_sold_in_multiple_quantities = "FALSE";
  listing.is_shipping_price_per_quantity_sold = "FALSE";

  listing.fpo_amount = "";
  listing.fpo_duration = "3 days";
  listing.fpo_to = "A";

  listing.av_bidders_only = "TRUE";
  listing.auction_length = "7 day";
  listing.auction_end_time = "";

  listing.delivery_pickup_allowed = "TRUE";
  listing.delivery_must_pickup = "FALSE";
  listing.delivery_use_bookcourier_rates = "FALSE";
  listing.delivery_bookcourier_is_box = "";
  listing.delivery_bookcourier_bag_size = "";
  listing.delivery_bookcourier_selected_courier = "";
  listing.delivery_bookcourier_service_level = "";
  listing.delivery_bookcourier_no_restricted_items = "";

  listing.delivery_price = getDeliveryPrice(listing);

  listing.payment_bank_deposit = "TRUE";
  listing.payment_credit_card = "TRUE";
  listing.payment_cash = "TRUE";
  listing.payment_afterpay = "TRUE";
  listing.payment_other = "";
  listing.send_payment_instructions = "TRUE";

  listing.youtube_video_key = "";
  listing.display_bold = "FALSE";
  listing.gallery = "FALSE";
  listing.gallery_plus = "FALSE";
  listing.feature = "FALSE";
  listing.super_feature = "FALSE";
  listing.donation_recipient = "";
  listing.folder = "";
  listing.exclude_shipping_promotion = "FALSE";
  listing.listing_footer_enabled = "TRUE";

  listing.height_cm = "0";
  listing.width_cm = "0";
  listing.length_cm = "0";
  listing.weight_kg = "0";

  listing.brand = "";
  listing.manufacturer_code = "";
  listing.barcode_gtin = "";
  listing.update_active_listings = "FALSE";

  if (listing.title && listing.product_id_for_member) {
    const idTag = `#${listing.product_id_for_member}`;

    if (!listing.title.includes(idTag)) {
      listing.title = `${listing.title} ${idTag}`;
    }
  }

  return listing;
}

function isJewelleryItem(listing, photoType = "") {
  const text = `${photoType} ${listing.title} ${listing.body}`.toLowerCase();

  return [
    "jewellery",
    "ring",
    "necklace",
    "pendant",
    "earrings",
    "bracelet",
    "bangle",
    "chain",
    "gold",
    "silver",
    "diamond"
  ].some(word => text.includes(word));
}

function isValuedJewellery(photoType = "") {
  return photoType.toLowerCase().includes("valued");
}

function formatValuedJewelleryBody(listing) {
  return `NEW RETAIL PRICE: ${listing.new_retail_price || ""}

MARKET VALUE: ${listing.market_value || ""}

HALLMARK/STAMPED: ${listing.hallmark || listing.metal || ""}

SIZE: ${listing.size || ""}

WEIGHT: ${listing.weight || ""}

STONES/DIAMOND: ${listing.stones || "PLEASE SEE VALUATION FOR MORE INFO"}`;
}

function formatStandardJewelleryBody(listing) {
  return `ITEM TYPE:
${listing.item_type || ""}

METAL:
${listing.metal || ""}

STONE:
${listing.stone || ""}

SIZE:
${listing.size || ""}

WEIGHT:
${listing.weight || ""}

COSMETIC CONDITION:
(${listing.condition || "B"})

(A) LIKE NEW
(B) GOOD CONDITION
(C) AVERAGE CONDITION
(D) POOR CONDITION`;
}

function formatFloorstockBody(listing) {
  return `MODEL:
${listing.model || ""}

ACCESSORIES:


COSMETIC CONDITION:
(${listing.condition || "B"})

(A) LIKE NEW
(B) GOOD CONDITION
(C) AVERAGE CONDITION
(D) POOR CONDITION`;
}

function applyListingTemplate(listing, photoType = "") {
  if (isJewelleryItem(listing, photoType) && isValuedJewellery(photoType)) {
    listing.folder = "Valued Jewellery";
    listing.delivery_price = "Free shipping";
    listing.body = formatValuedJewelleryBody(listing);
    return listing;
  }

  if (isJewelleryItem(listing, photoType)) {
    listing.folder = "Jewellery";
    listing.delivery_price = "Small bag";
    listing.body = formatStandardJewelleryBody(listing);
    return listing;
  }

  listing.folder = "Floorstock";
  listing.body = formatFloorstockBody(listing);
  return listing;
}

async function createListingFromPhotos(stockCode, photoFiles, photoType = "") {
  const photoIdList = photoFiles.join(";");
  const firstPhotoPath = path.join("uploads", photoFiles[0]);
  const imageBase64 = fs.readFileSync(firstPhotoPath, "base64");
  const categoryText = getCategoryText();

  const result = await openai.responses.create({
    model: "gpt-4.1-mini",

    input: [
      {
        role: "user",
        content: [
          {
            type: "input_text",
            text: `
Look at this item photo and create Trade Me CSV listing data.

Photo type: "${photoType}"

IMPORTANT RULES:

- Use "${stockCode}" as the product_id_for_member.
- Use "${photoIdList}" as the photo_id_list.
- attributes must always be blank.
- has_promo must always be "N".
- fpo_duration must always be "3 days".
- Leave ACCESSORIES blank for manual staff entry.
- Choose only ONE cosmetic condition letter: A, B, C, or D.
- Estimate a reasonable second-hand NZD market price based on the item photo, brand, model, age and condition.
- Set start_price and reserve_price to the same estimated amount.
- Set buy_now_price slightly higher than start_price.
- Use plain numbers only. Example: "149", not "$149".
- If unsure, choose a conservative second-hand price.

Use these category override rules FIRST:

${JSON.stringify(categoryOverrides, null, 2)}

If the item clearly matches an override keyword, use that override category_id.

Only if no override matches, choose the best category_id from this shortlist:

${categoryText}

The category_id must be either an override category_id or a Category Code from the shortlist.
Do not invent a category_id.
Do not leave category_id blank.

BODY FORMAT RULES:

If the item is jewellery, including rings, necklaces, pendants, earrings, bracelets, bangles, or chains, the body must use this exact format:

ITEM TYPE:
{ring/necklace/pendant/earrings/chain/bracelet/bangle/etc}

METAL:
{gold/silver/platinum/stainless steel/unknown}

STONE:
{stone type or none/unknown}

SIZE:
{ring size/chain length/etc if visible, otherwise leave blank}

WEIGHT:
{if visible/known, otherwise leave blank}

For all non-jewellery items, the body must use this exact format:

MODEL:
{detected_model_number}

ACCESSORIES:


COSMETIC CONDITION:
({best_condition_letter})

(A) LIKE NEW
(B) GOOD CONDITION
(C) AVERAGE CONDITION
(D) POOR CONDITION

Return ONLY valid JSON.
Do NOT use markdown.
Do NOT use backticks.

For valued jewellery:
- If photoType is "Valued Jewellery", extract the valuation details if visible.
- new_retail_price should be the valuation retail price.
- market_value should be the valuation/market value.
- hallmark should be the stamped metal mark, for example 9CT, 14K, 18CT.
- size should be ring size or chain length.
- weight should be item weight.
- stones should say "PLEASE SEE VALUATION FOR MORE INFO" unless stone details are clearly known.

Use this exact JSON structure:

{
  "product_id_for_member": "",
  "sku": "",
  "stock_amount": "1",
  "unlimited_stock": "FALSE",
  "category_id": "",
  "second_category_id": "",
  "title": "",
  "subtitle": "",
  "body": "",
  "is_new": "FALSE",
  "attributes": "",
  "is_legal_notice_read": "TRUE",
  "start_price": "",
  "reserve_price": "",
  "buy_now_price": "",
  "is_clearance": "FALSE",
  "was_price": "",
  "has_promo": "N",
  "is_sold_in_multiple_quantities": "FALSE",
  "is_shipping_price_per_quantity_sold": "FALSE",
  "fpo_amount": "",
  "fpo_duration": "3 days",
  "fpo_to": "A",
  "av_bidders_only": "TRUE",
  "auction_length": "7 day",
  "auction_end_time": "",
  "delivery_pickup_allowed": "TRUE",
  "delivery_must_pickup": "FALSE",
  "delivery_use_bookcourier_rates": "FALSE",
  "delivery_bookcourier_is_box": "",
  "delivery_bookcourier_bag_size": "",
  "delivery_bookcourier_selected_courier": "",
  "delivery_bookcourier_service_level": "",
  "delivery_bookcourier_no_restricted_items": "",
  "delivery_price": "",
  "payment_bank_deposit": "TRUE",
  "payment_credit_card": "TRUE",
  "payment_cash": "TRUE",
  "payment_afterpay": "TRUE",
  "payment_other": "",
  "send_payment_instructions": "TRUE",
  "photo_id_list": "",
  "youtube_video_key": "",
  "display_bold": "FALSE",
  "gallery": "FALSE",
  "gallery_plus": "FALSE",
  "feature": "FALSE",
  "super_feature": "FALSE",
  "donation_recipient": "",
  "folder": "",
  "exclude_shipping_promotion": "FALSE",
  "listing_footer_enabled": "TRUE",
  "height_cm": "0",
  "width_cm": "0",
  "length_cm": "0",
  "weight_kg": "0",
  "brand": "",
  "manufacturer_code": "",
  "barcode_gtin": "",
  "item_type": "",
  "metal": "",
  "stone": "",
  "stones": "",
  "size": "",
  "weight": "",
  "hallmark": "",
  "new_retail_price": "",
  "market_value": "",
  "model": "",	
  "condition": "",
  "update_active_listings": "FALSE"
}
`
          },
          {
            type: "input_image",
            image_url: `data:image/jpeg;base64,${imageBase64}`
          }
        ]
      }
    ]
  });

  const cleanListing = result.output_text
    .replace(/```json/g, "")
    .replace(/```/g, "")
    .trim();

  const listing = JSON.parse(cleanListing);

  listing.product_id_for_member = stockCode;
  listing.photo_id_list = photoIdList;

  applyTradeMeDefaults(listing);
return applyListingTemplate(listing, photoType);
}

app.post("/upload", upload.array("photos", 10), async (req, res) => {
  try {
    const uploadedFiles = req.files;

    if (!uploadedFiles || uploadedFiles.length === 0) {
      return res.status(400).json({
        error: "No photos uploaded"
      });
    }

    const firstFile = uploadedFiles[0];

    const stockCode = path
      .parse(firstFile.originalname)
      .name
      .replace(/\s*\(\d+\)$/g, "");

    const photoFiles = uploadedFiles.map(file => file.filename);

    const photoType = req.body.photoType || "Floorstock";

const listing = await createListingFromPhotos(stockCode, photoFiles, photoType);

    await csvWriter.writeRecords([listing]);

    res.json({
      message: "AI listing created",
      listing: JSON.stringify(listing)
    });

  } catch (error) {
    console.error(error);

    res.status(500).json({
      error: "Something went wrong"
    });
  }
});

app.get("/generate-from-uploads", async (req, res) => {
  try {
    const files = fs
      .readdirSync("uploads")
      .filter(file => /\.(jpg|jpeg|png)$/i.test(file));

    const groups = {};

    files.forEach(file => {
      const stockCode = path
        .parse(file)
        .name
        .replace(/\s*\(\d+\)$/g, "");

      if (!groups[stockCode]) {
        groups[stockCode] = [];
      }

      groups[stockCode].push(file);
    });

    const createdListings = [];

    for (const stockCode of Object.keys(groups)) {
      const photoFiles = groups[stockCode];

      const photoType = "Floorstock";

const listing = await createListingFromPhotos(
    stockCode,
    photoFiles,
    photoType
);

await csvWriter.writeRecords([listing]);

const processedFolder = path.join("uploads", "processed");

if (!fs.existsSync(processedFolder)) {
  fs.mkdirSync(processedFolder);
}

photoFiles.forEach(file => {
  const oldPath = path.join("uploads", file);
  const newPath = path.join(processedFolder, file);

  if (fs.existsSync(oldPath)) {
    fs.renameSync(oldPath, newPath);
  }
});

createdListings.push(listing);
    }

    res.json({
      success: true,
      message: "CSV rows created from uploads folder",
      count: createdListings.length,
      listings: createdListings
    });

  } catch (error) {
    console.error(error);

    res.status(500).json({
      success: false,
      error: "Failed to generate CSV rows from uploads folder"
    });
  }
});

app.get("/generate-from-dropbox-ready", async (req, res) => {
  try {
    const accessToken = await getDropboxAccessToken();

    const dbxReady = new Dropbox({
      accessToken
    });

    const readyFolder = "/Trademe CSV Queue/Ready";
    const processedFolder = "/Trademe CSV Queue/Processed";

const csvPath = "trade-me-auto-listings.csv";

if (fs.existsSync(csvPath)) {
  fs.unlinkSync(csvPath);
}

    const listResult = await dbxReady.filesListFolder({
      path: readyFolder,
      recursive: true
    });

console.log("Ready folder:", readyFolder);

console.log(
  "Raw Dropbox response:",
  JSON.stringify(listResult.result, null, 2)
);

console.log(
  "Dropbox entries:",
  listResult.result.entries.map(e => ({
    name: e.name,
    tag: e[".tag"],
    path: e.path_display
  }))
);

    const imageFiles = listResult.result.entries.filter(entry =>
      entry[".tag"] === "file" &&
      /\.(jpg|jpeg|png)$/i.test(entry.name)
    );

    if (imageFiles.length === 0) {
      return res.json({
        success: false,
        message: "No photos found in Dropbox Ready folder"
      });
    }

    const groups = {};

    for (const file of imageFiles) {
const stockCode = path
  .parse(file.name)
  .name
  .replace(/\s*\(\d+\)$/g, "");

      if (!groups[stockCode]) {
        groups[stockCode] = [];
      }

      groups[stockCode].push(file);
    }

    const createdListings = [];

    for (const stockCode of Object.keys(groups)) {
      const files = groups[stockCode];
      const localFiles = [];

      for (const file of files) {
        const download = await dbxReady.filesDownload({
          path: file.path_lower
        });

        const localFileName = file.name;
        const localPath = path.join("uploads", localFileName);

        fs.writeFileSync(localPath, download.result.fileBinary);

        localFiles.push(localFileName);
      }

      const photoType = "Floorstock";

const listing = await createListingFromPhotos(
    stockCode,
    localFiles,
    photoType
);

      await csvWriter.writeRecords([listing]);

      createdListings.push(listing);

      for (const file of files) {
        const processedPath = file.path_display.replace(
          readyFolder,
          processedFolder
        );

        await dbxReady.filesMoveV2({
          from_path: file.path_display,
          to_path: processedPath,
          autorename: true
        });
      }

      for (const localFile of localFiles) {
        const localPath = path.join("uploads", localFile);

        if (fs.existsSync(localPath)) {
          fs.unlinkSync(localPath);
        }
      }
    }

const csvContent = fs.readFileSync(csvPath);

await dbxReady.filesUpload({
  path: `/Trademe CSV Queue/Processed/CSV/trade-me-auto-listings-${Date.now()}.csv`,
  contents: csvContent,
  mode: "add"
});

    res.json({
      success: true,
      message: "CSV rows created from Dropbox Ready folder",
      count: createdListings.length,
      listings: createdListings
    });

  } catch (error) {
    console.error(error);

    res.status(500).json({
      success: false,
      error: "Failed to generate CSV from Dropbox Ready folder"
    });
  }
});

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});