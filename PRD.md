# Product Requirement Document

## Overview
Google Spreadsheet AppScript that fills missing data of spreadsheet serves as a vocaburary list.

Application Type: AppScript in Google Spreadsheet
License: MIT

## Features
1. Iterate all non-empty rows in the current sheet, and if the cell of "日本語" column is empty and the cell of "英語" is not empty, then translate the value in "英語" column and put the result in the "日本語" column.
2. Iterate all non-empty rows in the current sheet, and if the cell of "英語" column is empty and the cell of "日本語" is not empty, then translate the value in "日本語" column and put the result in the "英語" column.
3. Iterate all non-empty rows in the current sheet, and if the cell of "例文" column is empty and the cell of "英語" is not empty, then genrate a sentence as the usage example of the word/phrase in the cell of "英語" column, and put the result in the "例文" column.
4. Iterate all non-empty rows in the current sheet, and if the cell of "例文和訳" column is empty and the cell of "例文" is not empty, then translate the example sentence in the "例文" column to Japanese and put the result in the "例文和訳" column.
5. Iterate all non-empty rows in the current sheet, and if the cell of "類義語" column is empty and the cell of "英語" is not empty, then find synonyms of the word/phrase in the cell of "英語" column, and put the result in the "類義語" column.
6. For Features 3, 4, and 5, use the Japanese meaning (from "日本語" column) to help generate more accurate example sentences and synonyms that match the specific intended meaning of the English word/phrase.
7. For Japanese translations (Features 1 and 4), use plain, colloquial style such as "〜だ" or "〜である" for natural-sounding translations.

## Extra Requirements
- Execute each feature listed in "Features" section above in sequence, so that next feature will be use the data updated by the previous feature.
- Fill the cell with light purple, if the cell is updated by this script.

## Technical Requirments
- Use OpenAI API for translations.
- For each feature, use the value in the "文脈" column of the same row as the context of translation, if it is not empty.
- For example sentences and synonyms, use the Japanese translation to ensure the generated content reflects the specific meaning of the word/phrase as understood in the Japanese context.
