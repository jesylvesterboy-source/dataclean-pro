# 📊 DataClean Pro

**Automate the most painful part of data analysis — in one command.**

> Built by [**Eduxellence Analytics**](https://eduxellence.org) — Elite data solutions for research teams and enterprises worldwide.

---

## The Problem We Solved

Analysts and researchers spend **up to 80% of their project time** doing repetitive, manual data cleaning before a single model can be run. Messy CSVs with:

- Mixed date formats (`15/01/2024` vs `Jan 15 2024` vs `2024-01-15`)
- Currency symbols mixed into numeric columns (`$1,200.00` vs `1200`)
- Inconsistent categories (`MALE`, `male`, `Male ` — all the same thing)
- Hidden duplicates and impossible outlier values (`age: 999`)
- Missing values scattered across dozens of columns
- Blank rows and phantom null markers (`N/A`, `#N/A`, `-`, `.`, `None`)

**DataClean Pro fixes all of this automatically.**

---

## Demo: Before & After

**Input CSV (messy):**
```
respondent_id, age,  gender, cgpa, date_submitted
1,             23,   MALE,   3.75, 15/01/2024
2,             N/A,  female, ,     2024-01-15
3,             999,  Male,   3.9,  Jan 16 2024
2,             N/A,  female, ,     2024-01-15   ← duplicate!
```

**Output CSV (clean):**
```
respondent_id, age,  gender, cgpa, date_submitted
1,             23,   Male,   3.75, 2024-01-15
2,             23.5, Female, 3.7,  2024-01-15    ← age imputed (median), gender standardised
3,             23.5, Male,   3.9,  2024-01-16    ← outlier age replaced, date standardised
```

*Duplicate removed. 4 nulls imputed. 3 date formats → 1 standard. Zero manual work.*

---

## What It Does (11 Cleaning Stages)

| Stage | What Happens |
|-------|-------------|
| 🧹 Whitespace stripping | Cleans column names and all string cells |
| 🎯 Null standardisation | Converts `N/A`, `None`, `#N/A`, `-`, `.` → proper NaN |
| 🗑 Blank removal | Drops fully empty rows and columns |
| 🔍 Smart type detection | Auto-identifies numeric, date, categorical, text, boolean, identifier |
| 📝 Category standardisation | `MALE` / `male` / `Male ` → `Male` (Title Case) |
| 📅 Date standardisation | 8 date formats → single `YYYY-MM-DD` standard |
| 💰 Currency parsing | `$1,200.00` / `£850` / `₦45,000` → clean floats |
| 👥 Duplicate removal | Exact duplicate rows detected and dropped |
| ⚠️ Outlier flagging | IQR-based outlier detection — flagged in `_outlier_*` columns |
| 🩹 Missing imputation | Numerics → median fill · Categoricals → mode fill |
| 📋 Column intelligence | Full per-column audit: type, missing %, unique count, stats |

**Outputs:**
- ✅ Cleaned Excel file (`.xlsx`) with 4 sheets: Cleaned Data, Cleaning Log, Column Summary, Run Stats
- ✅ Beautiful standalone HTML audit report — shareable, professional, branded

---

## Quick Start — One Command

```bash
# 1. Install dependencies (one time only)
pip install -r requirements.txt

# 2. Clean a single CSV file
python dataclean_pro.py your_messy_file.csv

# 3. That's it. Check the output/ folder.
```

---

## All Usage Options

```bash
# Interactive mode (will show available CSVs and let you pick)
python dataclean_pro.py

# Clean a specific file
python dataclean_pro.py data/survey.csv

# Clean into a specific output folder
python dataclean_pro.py data/survey.csv --output results/

# Clean an ENTIRE folder of CSVs (batch mode)
python dataclean_pro.py --batch data/messy/

# Silent mode (no progress output — good for scripts/automation)
python dataclean_pro.py data/survey.csv --silent

# Adjust outlier sensitivity (default is 3.0 — higher = less aggressive)
python dataclean_pro.py data/survey.csv --outlier-factor 2.5
```

---

## Step-by-Step Setup Guide

### Step 1 — Check Python version

```bash
python --version
```
You need **Python 3.9 or higher**. Download from [python.org](https://python.org) if needed.

---

### Step 2 — Download DataClean Pro

**Option A — Git clone (recommended):**
```bash
git clone https://github.com/eduxellence/dataclean-pro.git
cd dataclean-pro
```

**Option B — Download ZIP:**
Click the green **Code** button above → **Download ZIP** → Extract the folder → Open terminal inside it.

---

### Step 3 — Install dependencies

```bash
pip install -r requirements.txt
```

This installs: `pandas`, `numpy`, `openpyxl`, `colorama`, `tabulate`, `scipy`.

> **Windows users:** If `pip` isn't found, try `python -m pip install -r requirements.txt`
> **Mac users:** If you have Python 3 installed via Homebrew, use `pip3` instead.

---

### Step 4 — Test with the included sample files

Three messy datasets are included in `data/messy/` so you can verify everything works immediately:

```bash
# Test 1: Academic survey data (mixed dates, null ages, duplicate rows)
python dataclean_pro.py data/messy/academic_survey.csv

# Test 2: Business sales data (currency symbols, mixed case, outliers)
python dataclean_pro.py data/messy/sales_data.csv

# Test 3: Financial research data (duplicates, blank rows, mixed nulls)
python dataclean_pro.py data/messy/financial_research.csv

# Test 4: Clean all three at once
python dataclean_pro.py --batch data/messy/
```

After running, open the `output/` folder. You will see:
- `*_cleaned_[timestamp].xlsx` — your cleaned data
- `*_report_[timestamp].html` — open this in any browser to see the full audit report

---

### Step 5 — Clean your own data

```bash
python dataclean_pro.py path/to/your/file.csv
```

Replace `path/to/your/file.csv` with your actual file path. Examples:
- `python dataclean_pro.py Downloads/survey_responses.csv`
- `python dataclean_pro.py C:\Users\John\Documents\sales_q1.csv` (Windows)
- `python dataclean_pro.py ~/Desktop/research_data.csv` (Mac/Linux)

---

## Understanding the HTML Report

Open the generated `*_report_*.html` file in Chrome, Firefox, or Edge. You will see:

| Section | What it Shows |
|---------|--------------|
| **Summary Cards** | Rows processed, rows removed, missing values before/after |
| **Cleaning Log** | Every action taken, with ✓ (ok), ⚠ (warning), or ✗ (error) |
| **Column Intelligence** | Type detected, % missing, unique count, min/max/mean |
| **Data Preview** | First 10 rows of the cleaned data |

---

## Understanding the Excel Output

| Sheet | Contents |
|-------|----------|
| **Cleaned Data** | Your analysis-ready dataset |
| **Cleaning Log** | Full audit trail of every operation |
| **Column Summary** | Per-column stats: type, missing %, sample values |
| **Run Stats** | Input rows, output rows, timestamp |

---

## Troubleshooting

**"No module named pandas"**
```bash
pip install pandas openpyxl numpy colorama tabulate
```

**"FileNotFoundError"**
Check your file path. Use the full path if you're getting errors:
```bash
python dataclean_pro.py "C:/Users/YourName/Desktop/data.csv"
```

**"UnicodeDecodeError"**
Your file has unusual characters. Open it in Excel → Save As → CSV UTF-8. Then run again.

**Output looks wrong / dates not parsed**
Your date format may be unusual. Open an issue and share a sample row — we'll add support.

---

## Requirements

```
Python >= 3.9
pandas >= 1.5.0
numpy >= 1.23.0
openpyxl >= 3.0.10
colorama >= 0.4.6
tabulate >= 0.9.0
```

No internet connection required after installation. Runs fully offline.

---

## Roadmap

- [ ] Google Colab notebook (no installation required)
- [ ] SPSS `.sav` file support
- [ ] Stata `.dta` file support
- [ ] R integration (output `.rds` directly)
- [ ] Custom cleaning rules via YAML config file
- [ ] GUI drag-and-drop interface (Tkinter)

---

## About Eduxellence

**DataClean Pro** is a free tool built by [**Eduxellence Analytics**](https://eduxellence.org) — a data solutions agency specialising in:

- 📊 Advanced statistical analysis (R, Python, SPSS, EViews)
- 🔬 Research data infrastructure for academic institutions
- 📈 Econometric modelling and financial forecasting
- 🏗 Custom data pipeline architecture for enterprises

👉 **Need something more powerful?** [Book a free 15-minute Data Strategy Audit](https://eduxellence.org/#contact) — we'll review your data stack and suggest exactly what would save your team the most time.

---

## Licence

MIT — free to use, modify, and distribute. Attribution appreciated.

---

<div align="center">

**Built with ❤️ by [Eduxellence Analytics](https://eduxellence.org)**

*Turning messy data into insight — one dataset at a time.*

[![Website](https://img.shields.io/badge/Website-eduxellence.org-blue?style=for-the-badge)](https://eduxellence.org)
[![Tool](https://img.shields.io/badge/Tool-DataClean%20Pro-green?style=for-the-badge)](https://eduxellence.org/free-tools)

</div>
