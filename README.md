# QFX Automation

This internal Python tool streamlines the processing of CSV files containing financial transaction data. Built with a custom Tkinter GUI, it transforms raw exports into clean, structured Excel workbooks ready for reporting and reconciliation.

## 🔧 What It Does

- Converts `.csv` financial data into a formatted `.xlsx` workbook
- Standardizes account names using company-specific abbreviations
- Formats date fields as `MM/DD/YYYY` and preserves them as text (Excel-safe)
- Generates unique transaction IDs for duplicate detection
- Splits data into three Excel sheets:
  - `Cleaned Data` – unique transactions
  - `Duplicates` – duplicate transaction IDs
  - `Double Duplicates` – recurring duplicates
- Removes sensitive fields like account numbers and check IDs
- Uses a lightweight, self-contained GUI for non-technical users

## 👨‍💻 Tech Highlights

- Python 3 with `pandas`, `openpyxl`, and `numpy`
- Tkinter for the desktop GUI
- File-safe operations for easy use by internal finance/admin teams
- Designed for Windows; no external deployment required

---

*Developed as part of internal automation efforts to reduce manual data handling and ensure consistent formatting across financial reports.*
