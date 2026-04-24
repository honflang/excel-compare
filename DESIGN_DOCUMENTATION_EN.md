# DataDiff Ops

## 1. Project Overview

This project is a Python-based Excel comparison tool that compares rows between two Excel files on the same sheet and generates a comparison result file. It supports configuring multiple columns as a composite unique key, selecting specific compare columns, and highlighting differing cells.

---

## 2. Methodology

### 2.1 Core Idea

- Use a composite unique business key to align rows between the two sheets.
- Compare only the configured columns, avoiding noisy full-table comparisons.
- Write comparison results into a copy of the main file, preserving original formatting and adding difference markers.

### 2.2 Data Flow

1. Read the config file `config.json`.
2. Copy the main file to create the result file, avoiding modification of the source.
3. Iterate each sheet according to configuration:
   - Locate the sheet by `index` or `name`;
   - Skip header or rows using `skip_lines`;
   - Extract `unique_columns` to build composite keys;
   - Extract `compare_columns` for value comparison.
4. Compare each row key between main and sub sheets:
   - If keys match and compare columns are equal: mark as "Matched";
   - If keys match but compare columns differ: mark as "Mismatched" and highlight cells;
   - If the row is missing in the sub file: mark "Missing in Sub";
   - If the row is missing in the main file: add a result row and mark "Missing in Main";
   - Handle duplicate keys and multiple matches with warnings.
5. Append summary statistics in the result file: total rows, matches, mismatches, missing rows, duplicates, etc.

### 2.3 Key Techniques

- `pandas` reads Excel into DataFrames for convenient column access.
- `openpyxl` writes the result file, adds cell fill and comments.
- `ColumnConfig.sub` supports substring extraction rules for composite keys or compare values. Current compare value processing is limited to substring extraction.
- The result file keeps the original main file structure and adds two columns: unique key and comparison result.

---

## 3. Features

### 3.1 Supported Features

- Multi-sheet comparison via the `sheets` array.
- Composite key matching using multiple columns.
- Selective compare columns.
- Difference highlighting: yellow fill and comments for mismatched cells.
- Summary stats: total rows, matched, mismatched, missing rows, duplicates, multiple matches.
- Configurable output directory via `output_path`.
- Supports `skip_lines` for custom header row handling.

### 3.2 Output Description

- The result file is named `compared-{main filename without extension}-{timestamp}.xlsx`.
- It is created by copying the main file, preserving format and column order.
- Two columns are inserted at the front of each processed sheet:
  - Column A: composite unique business key;
  - Column B: comparison result description.
- Summary statistics are written at the bottom of the result sheet, with color-coded values.

---

## 4. Configuration

Example config file:

```json
{
  "main_compare_file_path": "/path/to/main.xlsx",
  "sub_compare_file_path": "/path/to/sub.xlsx",
  "output_path": "/path/to/output",
  "skip_both": true,
  "sheets": [
    {
      "index": 1,
      "name": "Sheet1",
      "skip_lines": 0,
      "unique_columns": [
        {"name": "部门", "sub": [null, -2]},
        "系统"
      ],
      "compare_columns": [
        {"name": "状态1", "sub": [2]},
        {"name": "状态2", "sub": [1]}
      ]
    }
  ]
}
```

### 4.1 Main Fields

- `main_compare_file_path`: path to the main file.
- `sub_compare_file_path`: path to the sub file.
- `output_path`: output directory for the result file; if empty, uses current working directory.
- `skip_both`: when skipping rows, whether the sub sheet should also skip.
- `sheets`: list of sheet configurations.

### 4.2 Sheet Configuration

- `index`: sheet index starting at 0; if both `index` and `name` exist, `index` takes precedence.
- `name`: sheet name used when `index` is absent.
- `skip_lines`: number of rows to skip before comparing, default 0.
- `unique_columns`: list of columns for building unique keys. You can use either string or object form for each column definition.
- `compare_columns`: list of columns to compare. You can also use either string or object form.

### 4.3 Column Configuration

- `name`: column name.
- `sub`: optional substring rule:
  - `[n]`: take first n characters;
  - `[start, end]`: take the range;
  - `[null, -n]`: take the last n characters.

---

## 5. Limitations & Risks

### 5.1 Current Limitations

- Supports only Excel files, not CSV or other table formats.
- Comparison is string-based, which does support numbers, nulls, and case-sensitive text, but may not handle semantic equivalence for dates or numeric values across formats.
- Unique key matching uses Python list search, suitable for small to medium datasets but can be slow for large sheets.
- Duplicate keys or multiple matches are flagged, but no advanced duplicate-resolution logic is provided.
- Comparison only highlights row-level column differences and does not support complex rules or weighted comparisons.
- Result file is based on a copy of the main file, so rows added from the sub file may not preserve original source formatting.

### 5.2 Potential Improvements

- Use hash maps or DataFrame indexing to optimize key matching.
- Add support for CSV, ODS, and other file formats.
- Support more field preprocessing and comparison rules, such as numeric normalization, currency symbol cleaning, date format normalization, and text whitespace/case normalization.
- Support semantic type-aware comparisons (for example, date or numeric equivalence), rather than relying solely on string comparison.
- Add spell-check or fuzzy-match style auto-correction capabilities to better identify small textual variations in field names or data values.
- Offer flexible diff output formats such as diff summary tables or JSON/CSV export.
- Improve error handling with more detailed diagnostics.

---

## 6. How to Use

1. Install dependencies:
```bash
pip install pandas openpyxl
```
2. Edit `config.json`, set main file, sub file, and sheet configuration.
3. Run the script:
```bash
python main.py
```
or specify a config file:
```bash
python main.py my-config.json
```
4. Find the generated `compared-...xlsx` file in the `output_path` directory.
