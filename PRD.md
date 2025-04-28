# Product Requirement Document

## Overview
Google Spreadsheet AppScript that fills missing data of spreadsheet serves as a vocaburary list.

Application Type: AppScript in Google Spreadsheet
License: MIT

## Features
1. Iterate all non-empty rows in the current sheet, and if the cell of "日本語" column is empty and the cell of "英語" is not empty, then translate the value in "英語" column and put the result in the "日本語" column.
2. Iterate all non-empty rows in the current sheet, and if the cell of "英語" column is empty and the cell of "日本語" is not empty, then translate the value in "日本語" column and put the result in the "英語" column.
3. Iterate all non-empty rows in the current sheet, and if the cell of "例文" column is empty and the cell of "英語" is not empty, then genrate a sentence as the usage example of the word/phrase in the cell of "英語" column, and put the result in the "例文" column.
4. Iterate all non-empty rows in the current sheet, and if the cell of "類義語" column is empty and the cell of "英語" is not empty, then find synonyms of the word/phrase in the cell of "英語" column, and put the result in the "類義語" column.

## Extra Requirements
- Execute each feature listed in "Features" section above in sequence, so that next feature will be use the data updated by the previous feature.
- Fill the cell with light purple, if the cell is updated by this script.

## Technical Requirments
- Use OpenAI API for translations.
- For translation process, use the value in the "文脈" column of the same row as the context of translation, if it is not empty.
