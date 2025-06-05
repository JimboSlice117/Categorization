// --- CONFIGURATION ---
const OPENAI_MODEL = "gpt-4o-mini";
const OPENAI_TEMPERATURE = 0.0; 
const MAX_COMPLETION_TOKENS = 60; 
const MAX_CATEGORIES_TO_SEND = 75; 
const SHEET_NAME_PRODUCTS = "BaseTemplate";
const SHEET_NAME_CATEGORIES = "Shopify_Categories";
const TITLE_COLUMN = "A";
const DESCRIPTION_COLUMN = "T"; 
const VENDOR_COLUMN = "U";    // <<<< ADAPT THIS if your Vendor column is different
const OUTPUT_CATEGORY_COLUMN = "FY";
const MAX_API_RETRIES = 3; 

// --- USER TO POPULATE BASED ON CHATGPT ANALYSIS & THEIR DATA ---
const VENDOR_SPECIFIC_CATEGORY_PREFIXES = {
  "Pioneer DJ": [
    "DJ Equipment > DJ Controllers", "DJ Equipment > Mixers", "DJ Equipment > DJ Accessories",
    "DJ Equipment > DJ Cases > DJ Controller Cases", "Pro Audio > Loudspeakers", "Pro Audio > Headphones"
  ],
  "Mackie": [
    "Pro Audio > Loudspeakers", "Pro Audio > Mixers", "Pro Audio > Audio Consoles & Mixers > Audio Mixer Accessories",
    "Pro Audio > Live Sound & Portable Systems > PA Systems", "Pro Audio > Live Sound & Portable Systems > Power Amplifiers",
    "Pro Audio > Recording", "Pro Audio > Headphones"
  ],
  "Shure": ["Microphones > Wired Microphones", "Microphones > Wireless Microphones", "Microphones > Microphone Accessories", "Pro Audio > Headphones"],
  // Add many more vendor mappings here based on your ChatGPT analysis and data...
  "ExampleVendor": ["Specific Top Level > Specific SubCategory"]
};

const KEYWORD_TO_CATEGORY_GROUP_MAPPINGS = [
  {
    keywordsRegex: /speaker|loudspeaker|subwoofer|monitor\b|soundbar/i,
    targetGroupPrefixes: [
      "Pro Audio > Loudspeakers", "Home & Business AV > Home AV", "Business AV > Commercial Speakers",
      "Pro Audio > Live Sound & Portable Systems > PA Systems", "Pro Audio > Live Sound & Portable Systems > Wireless Speakers"
    ]
  },
  { keywordsRegex: /headphone|earbud|iem\b|earphones/i, targetGroupPrefixes: ["Pro Audio > Headphones", "Microphones > Wireless Microphones > Wireless IEM Systems"] },
  { keywordsRegex: /microphone|lavalier|shotgun|mic\b/i, targetGroupPrefixes: ["Microphones", "Pro Audio > Accessories > Microphone Stands", "Business AV > Audio Conferencing Systems > Conference Microphones"] },
  // ... Add more general keyword mappings ...
];

const TOP_LEVEL_CATEGORIES_FOR_FALLBACK = [ 
    "Business AV", "Computers & Peripherals", "DJ Equipment", "Home & Business AV",
    "Microphones", "Mobile", "Musical Instruments", "Pro Audio", "Pro Lighting", "Pro Video", "Used"
];
// --- END OF USER POPULATION SECTION ---


/**
 * Creates a custom menu when the Google Sheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Product Tools')
    .addItem('1. Set API Key', 'setApiKey') // MODIFIED Text and Function Name
    .addItem('2. Categorize Products (Fixed List)', 'categorizeProductsWithFixedList')
    .addToUi();
}

/**
 * Prompts the user to set and store their API key.
 */
function setApiKey() { // MODIFIED Function Name
  const ui = SpreadsheetApp.getUi();
  const userProperties = PropertiesService.getUserProperties();
  let result = ui.prompt('API Key Setup', 'Enter your API Key:', ui.ButtonSet.OK_CANCEL); // MODIFIED Text
  if (result.getSelectedButton() === ui.Button.OK && result.getResponseText()) {
    userProperties.setProperty('OPENAI_API_KEY', result.getResponseText().trim()); // Internal property name remains specific
    ui.alert('API Key saved for this user.'); // MODIFIED Text
  } else if (result.getSelectedButton() !== ui.Button.CANCEL) {
    ui.alert('API Key not set or invalid input.'); // MODIFIED Text
  }
}

/**
 * Pre-filters the master category list based on product vendor and keywords.
 */
function getRelevantCategoriesForProduct(productTitle, productDescription, productVendor, allMasterCategories) {
  const titleLower = productTitle ? String(productTitle).toLowerCase() : "";
  const descriptionLower = productDescription ? String(productDescription).toLowerCase() : "";
  const combinedText = titleLower + " " + descriptionLower;
  const vendorLower = productVendor ? String(productVendor).toLowerCase() : "";

  let candidateGroupPrefixes = new Set(); 
  let matchOccurred = false;

  for (const vendorKey in VENDOR_SPECIFIC_CATEGORY_PREFIXES) {
    if (vendorLower.includes(vendorKey.toLowerCase())) { 
      VENDOR_SPECIFIC_CATEGORY_PREFIXES[vendorKey].forEach(prefix => candidateGroupPrefixes.add(prefix));
      matchOccurred = true;
      Logger.log("Matched vendor '" + productVendor + "' to rule for: " + vendorKey);
      break; 
    }
  }

  if (!matchOccurred || candidateGroupPrefixes.size < 5) { 
    KEYWORD_TO_CATEGORY_GROUP_MAPPINGS.forEach(mapping => {
      if (mapping.keywordsRegex.test(combinedText)) {
        mapping.targetGroupPrefixes.forEach(prefix => candidateGroupPrefixes.add(prefix));
        matchOccurred = true; 
      }
    });
  }

  let filteredCategories = [];
  if (candidateGroupPrefixes.size > 0) {
    allMasterCategories.forEach(masterCategory => {
      for (const prefix of candidateGroupPrefixes) {
        if (masterCategory.startsWith(prefix)) {
          filteredCategories.push(masterCategory);
          break; 
        }
      }
    });
    filteredCategories = [...new Set(filteredCategories)]; 
  }

  if (matchOccurred && filteredCategories.length === 0) {
      Logger.log("Warning: Keyword/Vendor rules matched, but no categories found starting with prefixes: [" + Array.from(candidateGroupPrefixes).join(", ") + "] for '" + productTitle + "'. Broadening search.");
      const matchedTopLevels = new Set();
      candidateGroupPrefixes.forEach(prefix => {
          const topLevel = prefix.split(" > ")[0];
          if (TOP_LEVEL_CATEGORIES_FOR_FALLBACK.includes(topLevel)) {
              matchedTopLevels.add(topLevel);
          }
      });
      if (matchedTopLevels.size > 0) {
          allMasterCategories.forEach(masterCategory => {
              for (const topLevel of matchedTopLevels) {
                  if (masterCategory.startsWith(topLevel)) {
                      filteredCategories.push(masterCategory);
                      break;
                  }
              }
          });
          filteredCategories = [...new Set(filteredCategories)];
      }
  }

  if (filteredCategories.length === 0) {
    Logger.log("No categories found from vendor or keyword mapping for '" + productTitle + "'. Using general top-level fallback.");
    allMasterCategories.forEach(masterCategory => {
      for (const topLevel of TOP_LEVEL_CATEGORIES_FOR_FALLBACK) {
        if (masterCategory.startsWith(topLevel)) {
          filteredCategories.push(masterCategory); 
        }
      }
    });
    filteredCategories = [...new Set(filteredCategories)];
  }
  
  if (filteredCategories.length > MAX_CATEGORIES_TO_SEND) {
    Logger.log("Filtered list for '" + productTitle + "' has " + filteredCategories.length + " categories. Truncating to " + MAX_CATEGORIES_TO_SEND);
    filteredCategories = filteredCategories.slice(0, MAX_CATEGORIES_TO_SEND);
  }
  
  if (filteredCategories.length === 0 && allMasterCategories.length > 0) {
      Logger.log("CRITICAL FALLBACK: No categories selected by any filter for '" + productTitle + "'. Sending first 50 master categories to AI to prevent empty list.");
      return allMasterCategories.slice(0, 50); 
  }

  Logger.log("For product '" + productTitle + "' (Vendor: " + productVendor + "), sending " + filteredCategories.length + " relevant categories to AI: [" + filteredCategories.slice(0,5).join(", ") + (filteredCategories.length > 5 ? "..." : "") + "]");
  return filteredCategories;
}


/**
 * Processes product listings to categorize them using a fixed list and an external API.
 */
function categorizeProductsWithFixedList() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseTemplateSheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);
  const categoriesSheet = ss.getSheetByName(SHEET_NAME_CATEGORIES);

  if (!baseTemplateSheet) { ui.alert(`Error: Sheet "${SHEET_NAME_PRODUCTS}" not found.`); return; }
  if (!categoriesSheet) { ui.alert(`Error: Sheet "${SHEET_NAME_CATEGORIES}" not found.`); return; }

  const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY'); // Internal property name unchanged
  if (!apiKey) {
    ui.alert("Error: API Key not set. Please run '1. Set API Key' from the menu."); // MODIFIED Text
    return;
  }

  const allMasterCategories = categoriesSheet.getRange("A2:A" + categoriesSheet.getLastRow()).getValues().flat().filter(String);
  if (allMasterCategories.length === 0) {
    ui.alert("Error: No categories found in '" + SHEET_NAME_CATEGORIES + "' tab, column A (starting A2).");
    return;
  }
  Logger.log(`Loaded ${allMasterCategories.length} master categories.`);

  const lastRow = baseTemplateSheet.getLastRow();
  if (lastRow < 2) { ui.alert("No product data found in '" + SHEET_NAME_PRODUCTS + "'."); return; }

  const titles = baseTemplateSheet.getRange(TITLE_COLUMN + "2:" + TITLE_COLUMN + lastRow).getValues().flat();
  const descriptions = baseTemplateSheet.getRange(DESCRIPTION_COLUMN + "2:" + DESCRIPTION_COLUMN + lastRow).getValues().flat();
  const vendors = baseTemplateSheet.getRange(VENDOR_COLUMN + "2:" + VENDOR_COLUMN + lastRow).getValues().flat();
  const outputRange = baseTemplateSheet.getRange(OUTPUT_CATEGORY_COLUMN + "2:" + OUTPUT_CATEGORY_COLUMN + lastRow);
  const currentOutputValues = outputRange.getValues();

  const results = [];
  let processedCount = 0;

  for (let i = 0; i < titles.length; i++) {
    if (currentOutputValues[i][0] && !String(currentOutputValues[i][0]).startsWith("API_ERROR") && !String(currentOutputValues[i][0]).startsWith("SCRIPT_ERROR") && String(currentOutputValues[i][0]).toUpperCase() !== "NEEDS_MANUAL_REVIEW") {
      results.push([currentOutputValues[i][0]]);
      continue;
    }

    const title = titles[i];
    const description = descriptions[i];
    const vendor = vendors[i] || ""; 

    if (!title && !description) {
      results.push([""]); 
      continue;
    }

    Logger.log(`Processing row ${i + 2}: Title - "${title}", Vendor - "${vendor}"`);
    processedCount++;
    
    const categoriesForThisPrompt = getRelevantCategoriesForProduct(String(title), String(description), String(vendor), allMasterCategories);

    if (categoriesForThisPrompt.length === 0) {
        Logger.log("CRITICAL: categoriesForThisPrompt is empty for product: " + title + " even after fallbacks. Skipping API call, marking for review.");
        results.push(["NEEDS_MANUAL_REVIEW_NO_CATEGORIES_SENT"]);
        continue;
    }

    const categorizationPrompt = `You are an expert e-commerce product categorizer. Your task is to analyze the given product and select the single most appropriate category from the "AVAILABLE CATEGORIES" list provided below. This list is fixed and cannot be changed.

Instructions:
1.  **Detailed Product Analysis & Identification:**
    a.  **Core Attributes:** Scrutinize the Product Title and Description to thoroughly understand its specific function, type, and primary use.
    b.  **Brand-Specific Products:** Consider the vendor/brand when assessing the product type.
    c.  **Product Variations & Features:** Pay close attention to details indicating specific variations or key features.
2.  **Select the Single Most Appropriate Category from the Fixed "AVAILABLE CATEGORIES" List:**
    a.  **Review Available Options:** Carefully review the entire "AVAILABLE CATEGORIES" list you have been given for this product.
    b.  **Match Specificity (If Possible):** If a category in this list *closely and accurately* matches the specific product type, choose that specific category.
    c.  **Select Best Broader Category (If Specific Match Unavailable):** If no highly specific category exists in this list, you MUST select the broader category from this list that provides the most logical classification.
    d.  **Accessories:** If the item is an accessory, prioritize an accessory-type category from the list.
    e.  **Mandatory Selection:** You MUST select one category from the "AVAILABLE CATEGORIES" list for this product.
3.  **Final Output Requirements:**
    a.  Your final choice MUST be one single category name taken EXACTLY from the "AVAILABLE CATEGORIES" list. Do NOT invent or modify category names.
    b.  **Final Verification (Crucial): Before outputting your 'Chosen Category', mentally double-check: Is the category string you are about to provide an exact, character-for-character, case-sensitive copy of one single entry from the 'AVAILABLE CATEGORIES' list that was provided to you for this product? If it is not, you have made an error. You MUST go back, re-evaluate, and select an exact, verbatim match from that list.**
    c.  Your response must consist SOLELY of this chosen (and verified exact match) category name.

AVAILABLE CATEGORIES:
${categoriesForThisPrompt.join('\n')}

Product Title: "${title}"
Product Description: "${description}"

Chosen Category:`;

    let chosenCategory = "NEEDS_MANUAL_REVIEW_API_ISSUE"; 
    let attempts = 0;

    while (attempts < MAX_API_RETRIES) {
      const payload = {
        model: OPENAI_MODEL,
        messages: [{ role: "user", content: categorizationPrompt }],
        max_completion_tokens: MAX_COMPLETION_TOKENS,
        temperature: OPENAI_TEMPERATURE
      };
      const options = {
        'method': 'post',
        'headers': { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      try {
        const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode === 200) {
          const jsonResponse = JSON.parse(responseText);
          let extractedCategory = jsonResponse.choices && jsonResponse.choices[0] && jsonResponse.choices[0].message && jsonResponse.choices[0].message.content
                                  ? jsonResponse.choices[0].message.content.trim()
                                  : "";
          
          const lowerCaseAllowedCategories = categoriesForThisPrompt.map(cat => cat.toLowerCase());
          if (extractedCategory && lowerCaseAllowedCategories.includes(extractedCategory.toLowerCase())) {
            chosenCategory = categoriesForThisPrompt.find(cat => cat.toLowerCase() === extractedCategory.toLowerCase()) || extractedCategory;
          } else {
            Logger.log(`CRITICAL AI DEVIATION: AI returned category "${extractedCategory}" which is not in the (filtered) allowed list for title: "${title}". Filtered list had ${categoriesForThisPrompt.length} items. First item was: "${categoriesForThisPrompt.length > 0 ? categoriesForThisPrompt[0] : 'EMPTY_FILTERED_LIST'}". Assigning first from filtered list as fallback.`);
            chosenCategory = categoriesForThisPrompt.length > 0 ? categoriesForThisPrompt[0] : "NEEDS_MANUAL_REVIEW_AI_INVALID_CHOICE";
          }
          break; 

        } else if (responseCode === 429) { 
          attempts++;
          Logger.log(`Rate limit hit (attempt ${attempts}/${MAX_API_RETRIES}) for product: ${title}. Retrying after delay. Response: ${responseText}`);
          let retryAfterSeconds = 5 * attempts; 
          try {
            const errorPayload = JSON.parse(responseText);
            if (errorPayload && errorPayload.error && errorPayload.error.message) {
              const match = errorPayload.error.message.match(/Please try again in ([\d\.]+)s/);
              if (match && match[1]) {
                retryAfterSeconds = parseFloat(match[1]) + 0.5 + attempts; 
              }
            }
          } catch (e) { /* Ignore parsing error, use calculated retry */ }
          Utilities.sleep(Math.max(retryAfterSeconds * 1000, 2000)); 
          if (attempts >= MAX_API_RETRIES) {
            Logger.log(`Max retry attempts reached for product: ${title}.`);
            chosenCategory = `API_ERROR_RATE_LIMIT_MAX_RETRIES: ${responseCode}`;
            break;
          }
        } else { 
          Logger.log(`API Error for product ${title}: HTTP ${responseCode} - ${responseText}`);
          chosenCategory = `API_ERROR: ${responseCode}`;
          break; 
        }
      } catch (e) {
        Logger.log(`Script execution error during API call for product ${title}: ${e.toString()} ${e.stack}`);
        chosenCategory = `SCRIPT_ERROR: ${e.message}`;
        break; 
      }
    }
    results.push([chosenCategory]);
  }

  if (results.length > 0) {
    outputRange.setValues(results); 
    ui.alert(`Product categorization complete! ${processedCount} products were processed/re-processed in this run.`);
  } else {
    ui.alert("No products needed processing in this run or no products found.");
  }
}