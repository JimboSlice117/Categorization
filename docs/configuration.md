# Configuration constants

This project uses a single Apps Script file, `Categorization.js`, which contains several constants that control how the script interacts with a Google Sheet and the OpenAI API. The table below summarizes each constant, its role and typical values you might use when adapting the script.

| Constant | Purpose | Possible values |
|---------|---------|----------------|
|`OPENAI_MODEL`|Specifies which OpenAI chat model to use for categorization requests.|Any available chat completion model such as `gpt-4o-mini`, `gpt-4`, or `gpt-3.5-turbo`.|
|`OPENAI_TEMPERATURE`|Controls the randomness of the model’s output. A lower value makes responses more deterministic.|A number from `0` (fully deterministic) to around `2`.|
|`MAX_COMPLETION_TOKENS`|The maximum number of tokens the API may return for each request.|Typically between `1` and a few hundred depending on model limits and desired response length.|
|`MAX_CATEGORIES_TO_SEND`|Limits how many categories are passed to the model in a single prompt to reduce token usage.|Any positive integer; `75` is the default used here.|
|`SHEET_NAME_PRODUCTS`|Name of the sheet containing your product data.|Any valid sheet name, e.g. `BaseTemplate`.|
|`SHEET_NAME_CATEGORIES`|Name of the sheet that lists all categories available for assignment.|Any valid sheet name, such as `Shopify_Categories`.|
|`TITLE_COLUMN`|Column letter holding product titles.|Any single column letter like `A`.|
|`DESCRIPTION_COLUMN`|Column letter holding product descriptions.|Any single column letter, e.g. `T`.|
|`VENDOR_COLUMN`|Column letter containing vendor names.|Any single column letter; adjust if your sheet uses a different column than `U`.|
|`OUTPUT_CATEGORY_COLUMN`|Column where the chosen category will be written.|Any single column letter, here `FY`.|
|`MAX_API_RETRIES`|Number of times the script will retry an API call when it encounters rate limiting or temporary errors.|Any non-negative integer.|
|`VENDOR_SPECIFIC_CATEGORY_PREFIXES`|Mapping of vendor names to category prefixes that should be preferred when those vendors are detected.|Customize with your own vendor names and category paths.|
|`KEYWORD_TO_CATEGORY_GROUP_MAPPINGS`|Rules that map keywords to likely category prefixes. Used to pre-filter the master list before calling the API.|Customize each entry’s `keywordsRegex` and associated `targetGroupPrefixes`.|
|`TOP_LEVEL_CATEGORIES_FOR_FALLBACK`|Fallback list of high‑level categories used when no specific match can be found.|An array of strings containing your preferred top‑level categories.|

These constants are defined near the top of [`Categorization.js`](../Categorization.js) and can be adjusted to suit your specific spreadsheet layout or API preferences.
