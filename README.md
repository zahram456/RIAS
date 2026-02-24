# RIAS - Rehman Industries Accounting System

Desktop accounting application for Rehman Industries, built with Python, Tkinter, and SQLite.

![GitHub stars](https://img.shields.io/github/stars/zahram456/RIAS)

## Overview
RIAS is a voucher-driven double-entry accounting system for desktop use. It provides account management, voucher posting, reporting, export tools, database safety checks, and local backup support.

## Core Features
- Chart of accounts (`Asset`, `Liability`, `Income`, `Expense`)
- Voucher entry with debit/credit validation
- Posting controls to prevent edits on posted vouchers
- Voucher history with unbalanced voucher review tools
- Financial reports:
  - Trial Balance
  - Profit and Loss
  - Balance Sheet
  - General Ledger
  - Cash Flow
- Export options:
  - CSV export
  - PDF export (ReportLab)
  - Print-friendly PDF output
- Startup integrity and data-quality checks
- Automatic local database backups
- Theme settings (runtime switch from Settings screen)

## Requirements
- Python 3.9+
- Windows recommended (Tkinter desktop workflow)
- Dependency:
  - `reportlab>=3.6`

## Installation
1. Clone the repository:
   - `git clone https://github.com/zahram456/RIAS.git`
2. Enter the project directory:
   - `cd RIAS`
3. Create and activate a virtual environment:
   - `python -m venv .venv`
   - `.venv\Scripts\activate`
4. Install dependencies:
   - `pip install -r requirements.txt`

## Run
- Start the application:
  - `python rehman_accounting.py`

## Data Storage and Logs
RIAS stores runtime data in your home directory under `RIAS_Data`:
- Database:
  - `~/RIAS_Data/rehman_industries.db`
- Backups:
  - `~/RIAS_Data/db_backups/`
- Log file:
  - `~/RIAS_Data/rias.log`

## Keyboard Shortcuts
- Navigation:
  - `Ctrl+1` Accounts
  - `Ctrl+2` Voucher Entry
  - `Ctrl+3` Reports
  - `Ctrl+4` Voucher History
  - `Ctrl+5` Settings
- Voucher Entry:
  - `Ctrl+L` Add line
  - `Ctrl+S` Save voucher
  - `Ctrl+N` New voucher
- Reports:
  - `Alt+T` Trial Balance
  - `Alt+P` Profit and Loss
  - `Alt+B` Balance Sheet
  - `Alt+G` General Ledger
  - `Alt+C` Cash Flow
  - `Ctrl+E` Export CSV
  - `Ctrl+Shift+E` Export PDF
  - `Ctrl+P` Print

## Build (PyInstaller)
- Build from spec:
  - `pyinstaller RIAS.spec`
- Output:
  - `dist/`

## Notes
- If a bundled database exists and local data is missing, a starter database is seeded automatically.
- Report PDFs are generated using ReportLab.
