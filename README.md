# VBA Macro Repository 🚀

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)  
[![Contributions Welcome](https://img.shields.io/badge/Contributions-Welcome-brightgreen.svg)](CONTRIBUTING.md)  

Welcome to the **VBA Macro Repository**! This collection is dedicated to useful VBA macros that I have found really streamline my workflow and automate tedious tasks. More exciting macros will be added over time—stay tuned!

---

## Available Macros ✨

### 1. `toSqlValue` 📊

The `toSqlValue` macro is designed to convert tabular data into SQL value lines with smart content awareness. If you have multiple rows of data in Excel and need to insert them into a database, this macro will:

- **Wrap text values:** Automatically adds single quotes around text values.
- **Handle numeric values:** Leaves numbers unquoted.
- **Format rows:** Inserts the correct parentheses and commas.
- **Compile output:** Generates a single comma-separated string ready to paste into your terminal or DB viewer.

Example output:

```
('string 1', 5, 15, 'String 2'),
('string 3', 6, 12, 'string 4')
```

### 2. `DeleteEmptyRowsBasedOnFirstColumn` ✂️

The `DeleteEmptyRowsBasedOnFirstColumn` macro is designed to help you quickly clean up your Excel data by removing any rows where the first cell of your selected range is empty. For instance, consider the following data:

| String A  |           | 5  | 7  |
|-----------|-----------|----|----|
|           | String B  | 8  |    |
| String C  | String D  | 11 | 6  |

In this example, only the second row will be deleted because its first cell is blank.

**How It Works:**
- **Selection-Based:** Simply select the range you want to process.
- **Backward Looping:** The macro loops from the bottom row up to avoid skipping any rows during deletion.
- **Empty Check:** It uses the `Trim` function to ensure that even cells with only spaces are considered empty.

**Usage Instructions:**
1. Select the data range where you want to remove rows.
2. Run the macro.
3. The macro deletes any row within the selection that has an empty first cell.

### 3. `SelectDownToLastUsedCell` 📋

The `SelectDownToLastUsedCell` macro helps you quickly select all data in a column from your current position down to the last non-blank cell. This is particularly useful when working with large datasets where manually selecting ranges can be tedious.

**How It Works:**
- **Single-Cell Selection:** Start by selecting a single cell in any column.
- **Smart Detection:** The macro automatically finds the last used cell in that column.
- **Error Handling:** Includes checks for proper selection and handles cases where no data exists below the selected cell.

**Usage Instructions:**
1. Select exactly one cell in any worksheet.
2. Run the macro.
3. The selection will extend down to the last non-blank cell in that column.

### 4. `CompareTablesAndListDifferences` 🔍

The `CompareTablesAndListDifferences` macro compares two tables and identifies rows that exist in the first table but not in the second. This is useful for data reconciliation, finding missing entries, or validating data consistency across multiple tables.

**How It Works:**
- **Table Comparison:** Compares each row in the first table with all rows in the second table.
- **New Worksheet:** Creates a new worksheet named "Differences" to display the results.
- **Complete Row Output:** Shows the entire row content for entries that don't match.

**Usage Instructions:**
1. Set up two tables named "Table1" and "Table2" in sheets "Sheet1" and "Sheet2" respectively.
2. Run the macro.
3. Review the differences in the newly created "Differences" sheet.

### 5. `InsertConcatColumn` 🔗

The `InsertConcatColumn` macro creates a new column that concatenates the values from columns A and B with a space in between. This is helpful for quickly combining first and last names, addresses, or any other paired data points.

**How It Works:**
- **New Column:** Inserts a new column C right after column B.
- **Automatic Concatenation:** Combines values from columns A and B with a space separator.
- **Preserves Format:** The new column inherits formatting from adjacent columns.

**Usage Instructions:**
1. Ensure your data is in columns A and B.
2. Run the macro.
3. The concatenated results will appear in the newly created column C.

### 6. `copyRangeAsPicture` 📸

The `copyRangeAsPicture` macro allows you to quickly copy any selected range as an image that can be pasted into other applications. This is perfect for including Excel data in presentations, documents, or emails while preserving the exact visual appearance.

**How It Works:**
- **Interactive Selection:** Prompts you to select the range you want to copy.
- **Image Conversion:** Converts the selected range to a picture in the clipboard.
- **Confirmation:** Provides a message when the range has been successfully copied.

**Usage Instructions:**
1. Run the macro.
2. Select the range you want to copy as an image when prompted.
3. Paste the image into any application that accepts images.

### 7. `CreateTOC` 📑

The `CreateTOC` macro creates a table of contents sheet with hyperlinks to all other sheets in your workbook. This is especially useful for navigating large workbooks with many worksheets.

**How It Works:**
- **New Sheet:** Creates a new sheet named "Table of Contents" at the beginning of the workbook.
- **Hyperlinks:** Adds clickable hyperlinks to each worksheet.
- **Automatic Cleanup:** Removes any existing table of contents before creating a new one.

**Usage Instructions:**
1. Run the macro.
2. A new "Table of Contents" sheet will be created with links to all other sheets.
3. Click on any sheet name to navigate directly to that sheet.

### 8. `ExportSelectedSheetsToPDF` 📄

The `ExportSelectedSheetsToPDF` macro exports all selected sheets as individual PDF files. This is perfect for creating reports, archiving data, or sharing specific worksheets with others who don't need access to the entire workbook.

**How It Works:**
- **Multiple Sheet Support:** Exports each selected sheet as a separate PDF.
- **Automatic Folder Creation:** Creates a "PDF Export" folder in the same location as the workbook.
- **Named Output:** Each PDF is named after its corresponding worksheet.

**Usage Instructions:**
1. Select one or more sheets by holding Ctrl and clicking on sheet tabs.
2. Run the macro.
3. PDF files will be created in the "PDF Export" folder.

### 9. `HighlightDuplicateRows` 🎯

The `HighlightDuplicateRows` macro identifies and highlights duplicate rows in your data range, making it easy to spot redundant information. This is useful for data cleaning and ensuring data integrity.

**How It Works:**
- **Smart Detection:** Compares entire rows (columns A to C by default) to find duplicates.
- **Visual Highlighting:** Applies yellow highlighting to duplicate rows.
- **Preserves Original Data:** Only changes the formatting, not the data itself.

**Usage Instructions:**
1. Ensure your data starts from row 2 (row 1 is assumed to be headers).
2. Run the macro.
3. Duplicate rows will be highlighted in yellow.

### 10. `quickVersionBackup` 💾

The `quickVersionBackup` macro creates a timestamped copy of your current workbook. This is perfect for maintaining version history or creating quick backups before making significant changes.

**How It Works:**
- **Automatic Naming:** Creates a copy with the original filename plus a timestamp.
- **Same Location:** Saves the backup in the same folder as the original file.
- **Confirmation:** Displays a message with the path to the backup file.

**Usage Instructions:**
1. Run the macro.
2. A backup copy will be created with a timestamp in the filename.
3. Continue working with your original file.

---

## Future Additions 🔮

I'm continuously working on adding more useful VBA macros. If you have ideas or contributions, please feel free to reach out or submit a pull request!

---

## Contributing 🤝

Contributions are always welcome. If you'd like to contribute your own VBA macro or suggest improvements, please open an issue or submit a pull request. For guidelines, see [CONTRIBUTING.md](CONTRIBUTING.md).

---

## License 📜

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).
