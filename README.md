# Categorization

This project contains a Google Apps Script for automatically categorizing products in a spreadsheet using OpenAI.

## Spreadsheet Setup

1. **Products sheet** (defaults to `BaseTemplate`)
   - Column **A**: Product title
   - Column **T**: Product description
   - Column **U**: Vendor
   - Column **FY**: Output category
2. **Categories sheet** named `Shopify_Categories` with available categories listed in column **A**.

If the `Shopify_Categories` sheet is missing or empty, the script falls back to a large built-in list defined in `DEFAULT_CATEGORY_LIST` inside `Categorization.js`.

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

## Post-Processing Validation

Each categorization response must include a confidence level (High, Medium, or Low). Results that are not High are automatically flagged for manual review. The script also checks that vendor-specific categories match the configured `VENDOR_SPECIFIC_CATEGORY_PREFIXES`. If a mismatch occurs, the output is marked with an `AI_LOGIC_CONFLICT` error. Finally, nearly identical product titles categorized differently in the same run are flagged as duplicates for review.

