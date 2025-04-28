/**
 * Google Spreadsheet vocabulary helper
 * Automatically translates and fills missing data in vocabulary lists
 */

// Configuration
const MODEL = "gpt-4o"; // OpenAI model to use
const HIGHLIGHT_COLOR = "#E6E6FA"; // Light purple for highlighting updated cells

/**
 * Creates the menu item when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Vocabulary Helper')
    .addItem('Set OpenAI API Key', 'promptForApiKey')
    .addItem('Fill Missing Data', 'fillMissingData')
    .addToUi();
}

/**
 * Prompts the user for their OpenAI API key and saves it to script properties
 */
function promptForApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'OpenAI API Key',
    'Please enter your OpenAI API key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey) {
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('OPENAI_API_KEY', apiKey);
      ui.alert('API key saved successfully!');
    } else {
      ui.alert('Error: API key cannot be empty.');
    }
  }
}

/**
 * Gets the OpenAI API key from script properties or prompts the user if not set
 */
function getOpenAiApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let apiKey = scriptProperties.getProperty('OPENAI_API_KEY');
  
  if (!apiKey) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('OpenAI API key not found. Please set your API key first.');
    promptForApiKey();
    apiKey = scriptProperties.getProperty('OPENAI_API_KEY');
  }
  
  return apiKey;
}

/**
 * Main function to fill missing data in the spreadsheet
 */
function fillMissingData() {
  // Check for API key first
  const apiKey = getOpenAiApiKey();
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Operation cancelled: No API key provided.');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Find column indices
  let dataRange = sheet.getDataRange();
  let values = dataRange.getValues();
  const headers = values[0];
  const englishColIndex = headers.indexOf("英語");
  const japaneseColIndex = headers.indexOf("日本語");
  const exampleColIndex = headers.indexOf("例文");
  const synonymColIndex = headers.indexOf("類義語");
  const contextColIndex = headers.indexOf("文脈");
  
  // Check if required columns exist
  if (englishColIndex === -1 || japaneseColIndex === -1) {
    SpreadsheetApp.getUi().alert('Error: Required columns "英語" and/or "日本語" not found.');
    return;
  }
  
  // Process feature 1: English to Japanese translation
  processEnglishToJapanese(sheet, values, englishColIndex, japaneseColIndex, contextColIndex);
  
  // Refresh data after feature 1
  dataRange = sheet.getDataRange();
  values = dataRange.getValues();
  
  // Process feature 2: Japanese to English translation
  processJapaneseToEnglish(sheet, values, japaneseColIndex, englishColIndex, contextColIndex);
  
  // Refresh data after feature 2
  dataRange = sheet.getDataRange();
  values = dataRange.getValues();
  
  // Process feature 3: Generate example sentences
  if (exampleColIndex !== -1) {
    generateExampleSentences(sheet, values, englishColIndex, exampleColIndex);
    
    // Refresh data after feature 3
    dataRange = sheet.getDataRange();
    values = dataRange.getValues();
  }
  
  // Process feature 4: Generate synonyms
  if (synonymColIndex !== -1) {
    generateSynonyms(sheet, values, englishColIndex, synonymColIndex);
  }
  
  SpreadsheetApp.getUi().alert('Process completed!');
}

/**
 * Process English to Japanese translations
 */
function processEnglishToJapanese(sheet, values, englishColIndex, japaneseColIndex, contextColIndex) {
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Skip empty rows
    if (!row[englishColIndex]) continue;
    
    // Check if Japanese is empty and English is not
    if (!row[japaneseColIndex] && row[englishColIndex]) {
      // Get context if available
      const context = contextColIndex !== -1 ? row[contextColIndex] : "";
      
      // Translate
      const translation = translateText(row[englishColIndex], "en", "ja", context);
      
      // Update cell
      if (translation) {
        const cell = sheet.getRange(i + 1, japaneseColIndex + 1);
        cell.setValue(translation);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
    }
  }
}

/**
 * Process Japanese to English translations
 */
function processJapaneseToEnglish(sheet, values, japaneseColIndex, englishColIndex, contextColIndex) {
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Skip empty rows
    if (!row[japaneseColIndex]) continue;
    
    // Check if English is empty and Japanese is not
    if (!row[englishColIndex] && row[japaneseColIndex]) {
      // Get context if available
      const context = contextColIndex !== -1 ? row[contextColIndex] : "";
      
      // Translate
      const translation = translateText(row[japaneseColIndex], "ja", "en", context);
      
      // Update cell
      if (translation) {
        const cell = sheet.getRange(i + 1, englishColIndex + 1);
        cell.setValue(translation);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
    }
  }
}

/**
 * Generate example sentences
 */
function generateExampleSentences(sheet, values, englishColIndex, exampleColIndex) {
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Skip empty rows
    if (!row[englishColIndex]) continue;
    
    // Check if Example is empty and English is not
    if (!row[exampleColIndex] && row[englishColIndex]) {
      // Generate example sentence
      const example = generateExample(row[englishColIndex]);
      
      // Update cell
      if (example) {
        const cell = sheet.getRange(i + 1, exampleColIndex + 1);
        cell.setValue(example);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
    }
  }
}

/**
 * Find synonyms
 */
function generateSynonyms(sheet, values, englishColIndex, synonymColIndex) {
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Skip empty rows
    if (!row[englishColIndex]) continue;
    
    // Check if Synonym is empty and English is not
    if (!row[synonymColIndex] && row[englishColIndex]) {
      // Find synonyms
      const synonyms = generateSynonym(row[englishColIndex]);
      
      // Update cell
      if (synonyms) {
        const cell = sheet.getRange(i + 1, synonymColIndex + 1);
        cell.setValue(synonyms);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
    }
  }
}

/**
 * Translate text using OpenAI API
 */
function translateText(text, fromLang, toLang, context = "") {
  try {
    let prompt = "";
    if (context) {
      prompt = `Translate the following ${fromLang === "en" ? "English" : "Japanese"} text to ${toLang === "en" ? "English" : "Japanese"}. Context: ${context}\nText: ${text}`;
    } else {
      prompt = `Translate the following ${fromLang === "en" ? "English" : "Japanese"} text to ${toLang === "en" ? "English" : "Japanese"}: ${text}`;
    }
    
    const response = callOpenAI(prompt);
    return response.trim();
  } catch (e) {
    console.error("Translation error:", e);
    return "";
  }
}

/**
 * Generate example sentence using OpenAI API
 */
function generateExample(word) {
  try {
    const prompt = `Generate a natural and useful example sentence using the English word or phrase: "${word}". Only return the example sentence and nothing else.`;
    const response = callOpenAI(prompt);
    return response.trim();
  } catch (e) {
    console.error("Example generation error:", e);
    return "";
  }
}

/**
 * Generate synonyms using OpenAI API
 */
function generateSynonym(word) {
  try {
    const prompt = `Find 3-5 existing synonyms for the English word or phrase: "${word}". Return the synonyms as a comma-separated list. Only include well-established synonyms that would be found in a thesaurus.`;
    const response = callOpenAI(prompt);
    return response.trim();
  } catch (e) {
    console.error("Synonym finding error:", e);
    return "";
  }
}

/**
 * Call OpenAI API with the given prompt
 */
function callOpenAI(prompt) {
  const apiKey = getOpenAiApiKey();
  const url = "https://api.openai.com/v1/chat/completions";
  
  const payload = {
    model: MODEL,
    messages: [
      {role: "system", content: "You are a helpful assistant that specializes in language translation and vocabulary help."},
      {role: "user", content: prompt}
    ],
    temperature: 0.3,
    max_tokens: 200
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  
  if (json.error) {
    throw new Error(json.error.message);
  }
  
  return json.choices[0].message.content;
} 