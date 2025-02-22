Below is a stylized README with emojis and badges for your repository:

---

# VBA Macro Repository 🚀

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)  
[![Contributions Welcome](https://img.shields.io/badge/Contributions-Welcome-brightgreen.svg)](CONTRIBUTING.md)  

Welcome to the **VBA Macro Repository**! This collection is dedicated to useful VBA macros that I have found really streamline my workflow and automate tedious tasks. More exciting macros will be added over time—stay tuned!

---

## Available Macro ✨

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

This macro is a great tool to streamline your data cleaning process and ensure that your dataset only includes rows with valid entries.

---

## Future Additions 🔮

I'm continuously working on adding more useful VBA macros. If you have ideas or contributions, please feel free to reach out or submit a pull request!

---

## Contributing 🤝

Contributions are always welcome. If you'd like to contribute your own VBA macro or suggest improvements, please open an issue or submit a pull request. For guidelines, see [CONTRIBUTING.md](CONTRIBUTING.md).

---

## License 📜

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

