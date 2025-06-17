# 🔥 Nifty Scraper & Dashboard Automation 🔥

Welcome to the ultimate NIFTY 50 options scraper and dashboard generator! This repository hosts a Python script that:

* Fetches the live NIFTY option‑chain via the NSE API
* Filters for the current (front‑month) contracts
* Extracts StrikePrice, ExpiryDate, and implied volatilities for Calls (CE) and Puts (PE)
* Produces a timestamped Excel workbook with:

  * **Raw Data** (front‑month only)
  * **IV Data** (Strike, Expiry, Call/Put IV)
  * **Volatility Skew** chart embedded alongside

All wrapped in a one‑click workflow—bundle it into a standalone executable or run directly with Python. 📈

---

## 🚀 Features

* **Front‑Month Only**: Automatically selects the nearest expiry.
* **Timestamped Output**: Generates `nifty_dashboard_YYYYMMDD_HHMMSS.xlsx` each run.
* **Embedded Volatility Skew**: Matplotlib chart inserted into the IV Data sheet.
* **Portable**: Bundle with PyInstaller or run from source with a simple setup.

---

## 📦 Installation & Setup

### Clone the Repo

```bash
git clone https://github.com/ForceSensitiveSaiyan/nifty-scraper.git
cd nifty-scraper
```

### Create & Activate Virtual Environment (optional)

```bash
python -m venv venv
# macOS/Linux
source venv/bin/activate
# Windows
venv\Scripts\activate
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

---

## ⚙️ Usage

### Run from Source

```bash
python nifty_scraper.py
```

Each execution will produce a new Excel file named `nifty_dashboard_YYYYMMDD_HHMMSS.xlsx` in the current directory.

### Build a Standalone Executable

```bash
pip install pyinstaller
pyinstaller --onefile --name NiftyDashboard nifty_scraper.py
```

Copy `dist/NiftyDashboard.exe` anywhere—no Python required. Double‑click to run and generate the latest dashboard file.

---

## ⚖️ License

This project is licensed under the MIT License © \FSJ.

> *Crafted for traders and quants who demand the freshest vol data in one click.*
