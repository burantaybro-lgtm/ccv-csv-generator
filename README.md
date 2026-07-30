Project: CCV CSV Generator

Purpose
- Generate Trade Me CSVs from Dropbox photos.

Workflow
Dropbox Ready
↓
Download Images
↓
OpenAI Vision
↓
Generate Listing
↓
Write CSV
↓
Upload CSV
↓
Move Images to Processed

Future Features
- Multiple AI image analysis
- Jewellery valuation extraction
- Better pricing
- Automatic processing
- Dashboard
# CCV CSV Generator

The generator reads stock photos from Dropbox, creates Trade Me listing data,
uploads a CSV, and moves successfully processed photos.

## Jewellery reports

Create this Dropbox folder:

```text
/Trademe CSV Queue/Reports-Jewellery
```

For standard (unvalued) jewellery, upload both files using the same date:

```text
2026-07-30-unvalued-jewellery.xls
2026-07-30-jewellery-details.jpg
```

The Excel report supplies stock code, description, metal, stones and weight.
The screenshot supplies Size and CC Barcode. Multiple screenshots can be used:

```text
2026-07-30-jewellery-details-1.jpg
2026-07-30-jewellery-details-2.jpg
```

Valued jewellery photos and all pages of the Jewellery Valuers Company Ltd.
valuation belong in the Ready folder. Every filename must begin with the same
stock code as the item photo.

Items with missing or uncertain jewellery fields remain in Ready and appear in
the dashboard's Needs Review list.
