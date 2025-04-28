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
    generateExampleSentences(sheet, values, englishColIndex, exampleColIndex, contextColIndex);
    
    // Refresh data after feature 3
    dataRange = sheet.getDataRange();
    values = dataRange.getValues();
  }
  
  // Process feature 4: Generate synonyms
  if (synonymColIndex !== -1) {
    generateSynonyms(sheet, values, englishColIndex, synonymColIndex, contextColIndex);
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
      const translation = translateWithOpenAI(row[englishColIndex], "en", "ja", context);
      
      // Update cell
      if (translation) {
        const cell = sheet.getRange(i + 1, japaneseColIndex + 1);
        cell.setValue(translation);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
      
      // Add a short delay to avoid hitting API rate limits
      Utilities.sleep(200);
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
      const translation = translateWithOpenAI(row[japaneseColIndex], "ja", "en", context);
      
      // Update cell
      if (translation) {
        const cell = sheet.getRange(i + 1, englishColIndex + 1);
        cell.setValue(translation);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
      
      // Add a short delay to avoid hitting API rate limits
      Utilities.sleep(200);
    }
  }
}

/**
 * Generate example sentences
 */
function generateExampleSentences(sheet, values, englishColIndex, exampleColIndex, contextColIndex) {
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Skip empty rows
    if (!row[englishColIndex]) continue;
    
    // Check if Example is empty and English is not
    if (!row[exampleColIndex] && row[englishColIndex]) {
      // Get context if available
      const context = contextColIndex !== -1 ? row[contextColIndex] : "";
      
      // Generate example sentence
      const example = generateExampleWithOpenAI(row[englishColIndex], context);
      
      // Update cell
      if (example) {
        const cell = sheet.getRange(i + 1, exampleColIndex + 1);
        cell.setValue(example);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
      
      // Add a short delay to avoid hitting API rate limits
      Utilities.sleep(200);
    }
  }
}

/**
 * Find synonyms
 */
function generateSynonyms(sheet, values, englishColIndex, synonymColIndex, contextColIndex) {
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Skip empty rows
    if (!row[englishColIndex]) continue;
    
    // Check if Synonym is empty and English is not
    if (!row[synonymColIndex] && row[englishColIndex]) {
      // Get context if available
      const context = contextColIndex !== -1 ? row[contextColIndex] : "";
      
      // Find synonyms
      const synonyms = findSynonymsWithOpenAI(row[englishColIndex], context);
      
      // Update cell
      if (synonyms) {
        const cell = sheet.getRange(i + 1, synonymColIndex + 1);
        cell.setValue(synonyms);
        cell.setBackground(HIGHLIGHT_COLOR);
      }
      
      // Add a short delay to avoid hitting API rate limits
      Utilities.sleep(200);
    }
  }
}

/**
 * Makes a request to the OpenAI API
 * 
 * @param {object} requestData - The data to send to the API
 * @return {object} The API response
 */
function callOpenAI(requestData) {
  const apiKey = getOpenAiApiKey();
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(requestData),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error('OpenAI API error: ' + error);
    throw new Error('Error calling OpenAI API: ' + error);
  }
}

/**
 * Translates text from source language to target language using OpenAI
 * 
 * @param {string} text - The text to translate
 * @param {string} sourceLang - The source language code
 * @param {string} targetLang - The target language code
 * @param {string} context - Optional context for translation
 * @return {string} The translated text
 */
function translateWithOpenAI(text, sourceLang, targetLang, context = "") {
  try {
    // Map language codes to full names for better OpenAI understanding
    const languageMap = {
      'en': 'English',
      'ja': 'Japanese'
    };
    
    const sourceLanguage = languageMap[sourceLang] || sourceLang;
    const targetLanguage = languageMap[targetLang] || targetLang;
    
    let systemPrompt = `You are a professional translator specialized in ${sourceLanguage} and ${targetLanguage}. Translate the given text from ${sourceLanguage} to ${targetLanguage}. Provide ONLY the translation, without any explanations or additional text.`;
    
    if (context) {
      systemPrompt += ` Use the following context to inform your translation: ${context}`;
    }
    
    const requestData = {
      'model': MODEL,
      'messages': [
        {
          'role': 'system',
          'content': systemPrompt
        },
        {
          'role': 'user',
          'content': text
        }
      ],
      'temperature': 0.3,
      'max_tokens': 500
    };
    
    const response = callOpenAI(requestData);
    if (response && response.choices && response.choices.length > 0) {
      return response.choices[0].message.content.trim();
    } else {
      console.error('Unexpected OpenAI response format for translation:', JSON.stringify(response));
      return targetLang === 'ja' ? '翻訳エラー' : 'Translation error';
    }
  } catch (error) {
    console.error('Translation error: ' + error);
    return targetLang === 'ja' ? '翻訳エラー' : 'Translation error';
  }
}

/**
 * Generate example sentence using OpenAI API
 * 
 * @param {string} word - The English word or phrase to generate an example for
 * @param {string} context - Optional context for example generation
 * @return {string} The example sentence
 */
function generateExampleWithOpenAI(word, context = "") {
  try {
    let systemPrompt = 'You are a language teacher who creates clear, natural example sentences using English vocabulary. Create a sentence that demonstrates the correct usage of the given word or phrase. Return ONLY the example sentence and nothing else.';
    
    if (context) {
      systemPrompt += ` Consider this context when creating the example: ${context}`;
    }
    
    const requestData = {
      'model': MODEL,
      'messages': [
        {
          'role': 'system',
          'content': systemPrompt
        },
        {
          'role': 'user',
          'content': word
        }
      ],
      'temperature': 0.7,
      'max_tokens': 150
    };
    
    const response = callOpenAI(requestData);
    if (response && response.choices && response.choices.length > 0) {
      return response.choices[0].message.content.trim();
    } else {
      console.error('Unexpected OpenAI response format for example generation:', JSON.stringify(response));
      return '例文生成エラー'; // Example generation error in Japanese
    }
  } catch (error) {
    console.error('Example generation error: ' + error);
    return '例文生成エラー'; // Example generation error in Japanese
  }
}

/**
 * Find synonyms using OpenAI API
 * 
 * @param {string} word - The English word or phrase to find synonyms for
 * @param {string} context - Optional context for synonym finding
 * @return {string} A comma-separated list of synonyms
 */
function findSynonymsWithOpenAI(word, context = "") {
  try {
    let systemPrompt = 'You are a language expert who identifies accurate synonyms for English words and phrases. Find 3-5 existing synonyms for the given word or phrase. Return ONLY a comma-separated list of synonyms, with no other text or explanations. Only include well-established synonyms that would be found in a thesaurus.';
    
    if (context) {
      systemPrompt += ` Use this context to find the most appropriate synonyms: ${context}`;
    }
    
    const requestData = {
      'model': MODEL,
      'messages': [
        {
          'role': 'system',
          'content': systemPrompt
        },
        {
          'role': 'user',
          'content': word
        }
      ],
      'temperature': 0.3,
      'max_tokens': 100
    };
    
    const response = callOpenAI(requestData);
    if (response && response.choices && response.choices.length > 0) {
      return response.choices[0].message.content.trim();
    } else {
      console.error('Unexpected OpenAI response format for synonym finding:', JSON.stringify(response));
      return '類義語検索エラー'; // Synonym finding error in Japanese
    }
  } catch (error) {
    console.error('Synonym finding error: ' + error);
    return '類義語検索エラー'; // Synonym finding error in Japanese
  }
} 