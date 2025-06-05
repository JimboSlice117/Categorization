# Categorization

This project contains a Google Apps Script for automatically categorizing products in a spreadsheet using OpenAI.

## Spreadsheet Setup

1. **Products sheet** (defaults to `BaseTemplate`)
   - Column **A**: Product title
   - Column **T**: Product description
   - Column **U**: Vendor
   - Column **FY**: Output category
2. **Categories sheet** named `Shopify_Categories` with available categories listed in column **A**.

## Usage

1. Open the Google Sheet and choose **Extensions → Apps Script**.
2. Replace the contents of the script editor with `Categorization.js` from this repository.
3. Review the configuration constants at the top of the script and modify them if your sheet uses different columns or sheet names.
4. Reload the sheet. A new **Product Tools** menu will appear.
5. Select **1. Set API Key** and enter your OpenAI key.
6. Run **2. Categorize Products (Fixed List)** to categorize all rows that do not yet have a value in the output column.
7. To process very large sheets incrementally, use **3. Categorize Next Batch** which processes up to the configured batch size each run.

## Customization

Vendor‑specific prefixes and keyword mappings used to filter the category list can be edited in the `VENDOR_SPECIFIC_CATEGORY_PREFIXES` and `KEYWORD_TO_CATEGORY_GROUP_MAPPINGS` objects inside `Categorization.js`. Adjust these arrays to influence how categories are suggested for your products.
The `BATCH_SIZE` constant controls how many rows are processed when running the batch command.

