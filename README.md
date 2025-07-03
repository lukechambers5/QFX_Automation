# QFX Automation

This internal Python tool streamlines the processing of CSV files containing financial transaction data. Built with a custom Tkinter GUI, it transforms raw exports into clean, structured Excel workbooks ready for reporting and reconciliation.

## What It Does

- Accepts `.csv` files with transactional financial data
- Cleans and standardizes column values:
  - Abbreviates `ACCOUNT_NAME` for easier reading
  - Deletes unwanted columns (e.g., `DFI_ID`, `FITID`)
- Converts `DTPOSTED` to readable `DATE` column
- Generates a unique `ID` for each row based on:
  - `ACCOUNT_NAME` + Excel-style serial date + `TRNAMT`
- Detects and splits:
  - Duplicates → saved in a separate sheet
  - Double duplicates → isolated to a third sheet
- Exports cleaned results to an `.xlsx` file with multiple tabs
- Preserves date formatting for Excel compatibility
- Provides a modern and intuitive GUI with file selection and open buttons

---

## GUI Features

- Built with `tkinter` and styled using `tkFont`
- Drag-to-move custom window behavior
- Interactive button hover effects
- Inline progress status label (`Processing...`, `Complete!`, `Failed`)
- One-click to open the generated Excel file

---

## Tech Stack

- **Python 3.10+**
- `tkinter` (built-in GUI library)
- `pandas` for data manipulation
- `openpyxl` for Excel export
- `numpy` for data type handling

---

## Input File Requirements

- Must be a `.csv` file
- Required columns:
  - `ACCOUNT_NAME`
  - `DTPOSTED` (YYYYMMDD format)
  - `TRNAMT` (transaction amount)
  - `NAME` (for additional duplicate ID clarification)

---

## Output

- Processed file is saved in the same directory as the original with the format:


---

## ⚠️ Note
- Internal Use Only!
- This tool is intended solely for authorized staff at Blitz Medical Billing. 
- Do **not** distribute or use this application outside approved environments.

## 🔐 Security & Compliance

This application is designed to process healthcare data (e.g., names, MRNs, dates of service), but no PHI is committed to or stored in this repository.

- No patient data is hardcoded or bundled.
- All uploads happen locally on the user's machine.
- Temporary and output files are not uploaded or retained externally.
- Users are responsible for ensuring HIPAA compliance when operating this tool in a production environment.

This repository contains logic only and is safe for internal, private use.