# RIAS – Desktop Accounting System

![GitHub stars](https://img.shields.io/github/stars/zahram456/RIAS)

Desktop accounting system for Rehman Industries built with Python, Tkinter, and SQLite.

**Features**
- Chart of accounts with account types: Asset, Liability, Income, Expense
- Voucher-based double-entry transactions (debit/credit validation)
- Ledger and reporting views
- CSV export utilities
- PDF report generation (ReportLab)
- Automatic local database backups

**Requirements**
- Python
- Windows recommended for the packaged build

**Quick Start**
1. Create a virtual environment and install dependencies:
   - `python -m venv .venv`
   - `.venv\Scripts\activate`
   - `pip install -r requirements.txt`
2. Run the app:
   - `python rehman_accounting.py`

**Data Location**
- The database is stored under your home directory in `RIAS_Data`:
  - `~/RIAS_Data/rehman_industries.db`
- Backups are stored in:
  - `~/RIAS_Data/db_backups/`
- Logs are written to:
  - `~/RIAS_Data/rias.log`

**Build (PyInstaller)**
- Build using the provided spec file:
  - `pyinstaller RIAS.spec`
- Output is written to `dist/`.

## 🛠 Installation
1. Clone the repo  
   `git clone https://github.com/zahram456/RIAS.git`

2. Create virtual environment  
   `python -m venv .venv`

3. Activate and install  
   `pip install -r requirements.txt`

4. Run  
   `python rehman_accounting.py`

**Notes**
- The app seeds a starter database if a bundled database is present.
- PDF reports are generated with ReportLab.
