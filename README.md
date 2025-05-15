# Google Sheets Vocabulary Helper

A Google Apps Script tool to automatically fill missing data in vocabulary lists.

## Features

- Translates English words to Japanese
- Translates Japanese words to English
- Generates example sentences for English words
- Finds synonyms for English words
- Highlights updated cells with light purple color
- Uses OpenAI API for translation and content generation
- Securely stores API key in script properties
- Uses Japanese meanings to generate more accurate examples and synonyms

## Setup Instructions

1. Open your Google Spreadsheet
2. Go to Extensions > Apps Script
3. Copy and paste the `Code.gs` file content into the Apps Script editor
4. Save the project
5. Refresh your spreadsheet

## Usage

1. Ensure your spreadsheet has the following columns:
   - "英語" (English)
   - "日本語" (Japanese)
   - "例文" (Example sentences)
   - "類義語" (Synonyms)
   - "文脈" (Context) - optional, provides context for all operations

2. Click on the "Vocabulary Helper" menu item that appears in your spreadsheet
3. Select "Set OpenAI API Key" to securely save your API key
4. Select "Fill Missing Data" to process all rows and fill in missing data

## Notes

- The script processes each feature in sequence
- Each feature uses the updated data from previous features
- Cells updated by the script are highlighted in light purple
- If the "文脈" column has a value, it will be used as context for all operations (translations, examples, and synonyms)
- The Japanese meaning of each word is used when generating example sentences and synonyms to ensure they match the specific intended meaning
- Your API key is stored securely in script properties, not in the code

## License

MIT 