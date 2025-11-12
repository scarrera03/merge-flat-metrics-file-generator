# merger.py

`merger.py` is a command-line utility that reads Excel workbooks — including multi-sheet files where each sheet represents a company — and produces a **flat CSV file** that includes only the columns from **"Revenue"** through **"Cash Flow"** (inclusive).  
It automatically skips non-relevant sections such as “Includes Depreciation”.

---

## Features

- Accepts an Excel file as input (`.xlsx`).
- Handles multi-sheet company workbooks or already-flattened sheets.
- Detects and normalizes headers automatically (case- and whitespace-insensitive).
- Keeps all columns from **Revenue → Cash Flow**, inclusive.
- Outputs a CSV file with the same base name as the input file.
- Provides optional flags:
  - `--sheet` → specify a sheet name to process directly.
  - `--list-sheets` → list all available sheets in the workbook.

---

## Installation

1. Clone or copy this project to your computer.
2. Create and activate a Python environment (recommended):
   ```bash
   conda create -n merger310 python=3.10 -y
   conda activate merger310
   ```
3. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   or directly:
   ```bash
   pip install pandas openpyxl
   ```

---

## Usage

### Basic command
```bash
python3 merger.py -i "Data for File Combining Test.xlsx"
```

This reads your Excel file and generates:
```
Data for File Combining Test.csv
```

### Specify a sheet
If your workbook has multiple sheets and you want to process just one:
```bash
python3 merger.py -i "Data for File Combining Test.xlsx" --sheet "Flat"
```

### List available sheets
```bash
python3 merger.py -i "Data for File Combining Test.xlsx" --list-sheets
```

---

## How It Works

1. **Input reading**  
   The program loads your Excel workbook and automatically identifies:
   - Sheets that contain yearly views (e.g., `2024B`, `2024F`).
   - Or an already flattened sheet with headers including “Revenue” and “Cash Flow”.

2. **Data extraction**  
   - It scans headers, normalizes them (trims spaces, ignores case).
   - Finds the position of the **Revenue** and **Cash Flow** columns.
   - Keeps all columns from **Revenue** to **Cash Flow**, inclusive, plus any identifying columns that precede them (e.g., Company Name, Year, Month).

3. **Filtering**  
   - Automatically skips “Includes Depreciation”.
   - Keeps only December data (fixed month, as required).

4. **Output generation**  
   The filtered data is saved to a `.csv` file with the same base name as the input Excel file.

---

## Example

**Input:**  
`Data for File Combining Test.xlsx`

| Company Name | Year | Month | Version | View | Revenue | Cost of Goods Sold1 | Gross Margin | Cash Flow | Plus: Depreciation |
|---------------|------|--------|----------|------|----------|---------------------|---------------|------------|--------------------|
| Company A | 2024 | December | B | 2024B | 500000 | 420000 | 80000 | 65000 | 1000 |

**Output (CSV):**

| Company Name | Year | Month | Version | View | Revenue | Cost of Goods Sold1 | Gross Margin | Cash Flow |
|---------------|------|--------|----------|------|----------|---------------------|---------------|------------|
| Company A | 2024 | December | B | 2024B | 500000 | 420000 | 80000 | 65000 |

---

## Help

To see usage details:
```bash
python3 merger.py -h
```

**Output:**
```
usage: merger.py -i <input_excel_file> [--sheet SHEET_NAME] [--list-sheets]

Extract columns from 'Revenue' to 'Cash Flow' (inclusive) from an Excel file and save as CSV.

options:
  -h, --help       show this help message and exit
  -i, --input      Path to the Excel file to process (e.g., data.xlsx)
  --sheet          Optional sheet name to read
  --list-sheets    List sheet names and exit
```

---

## License

This project is open source and free to use for educational or business automation purposes.
