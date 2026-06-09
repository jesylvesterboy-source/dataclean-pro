#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║            DataClean Pro — by Eduxellence Analytics              ║
║              https://eduxellence.org | Free Tool v1.0            ║
╚══════════════════════════════════════════════════════════════════╝

Automates the entire CSV/Excel data cleaning & validation pipeline.
Handles: Duplicates · Missing Values · Outliers · Type Inference ·
         Date Standardisation · Currency Parsing · HTML Report Output

Usage:
    python dataclean_pro.py                        # interactive mode
    python dataclean_pro.py mydata.csv             # direct CSV file
    python dataclean_pro.py mydata.xlsx            # direct Excel file
    python dataclean_pro.py mydata.csv --silent    # no prompts
    python dataclean_pro.py --batch data/          # clean entire folder
    python dataclean_pro.py --gui                  # launch GUI

Google Colab (one-click):
    See dataclean_pro_colab.ipynb  — or paste the COLAB CELL below
    into any notebook and run it.
"""

import os
import sys
import csv
import re
import json
import time
import shutil
import argparse
import warnings
from pathlib import Path
from datetime import datetime
from copy import deepcopy

import numpy as np
import pandas as pd
from tabulate import tabulate

warnings.filterwarnings("ignore")

# ─── Colour helpers ────────────────────────────────────────────────────
try:
    from colorama import init, Fore, Style
    init(autoreset=True)
    C = {
        "header": Fore.CYAN + Style.BRIGHT,
        "ok":     Fore.GREEN + Style.BRIGHT,
        "warn":   Fore.YELLOW,
        "err":    Fore.RED + Style.BRIGHT,
        "dim":    Style.DIM,
        "reset":  Style.RESET_ALL,
        "blue":   Fore.BLUE + Style.BRIGHT,
        "magenta":Fore.MAGENTA + Style.BRIGHT,
    }
except ImportError:
    C = {k: "" for k in ("header","ok","warn","err","dim","reset","blue","magenta")}


# ─── Branding ────────────────────────────────────────────────────────────────
BANNER = f"""
{C['header']}╔══════════════════════════════════════════════════════════════════╗
║            DataClean Pro — by Eduxellence Analytics              ║
║              https://eduxellence.org | Free Tool v1.0            ║
╚══════════════════════════════════════════════════════════════════╝{C['reset']}
"""

BRAND = "Eduxellence Analytics · https://eduxellence.org"


# ─── Progress bar ──────────────────────────────────────────────────────────────
def progress_bar(label: str, total: int, current: int, width: int = 30) -> None:
    filled = int(width * current / max(total, 1))
    bar = "█" * filled + "░" * (width - filled)
    pct = int(100 * current / max(total, 1))
    print(f"\r  {C['dim']}{label}{C['reset']} [{C['ok']}{bar}{C['reset']}] {pct}%", end="", flush=True)
    if current >= total:
        print()


def step_banner(text: str) -> None:
    print(f"\n{C['blue']}▶  {text}{C['reset']}")


def ok(text: str) -> None:
    print(f"  {C['ok']}✓{C['reset']}  {text}")


def warn(text: str) -> None:
    print(f"  {C['warn']}⚠{C['reset']}  {text}")


def err(text: str) -> None:
    print(f"  {C['err']}✗{C['reset']}  {text}")


# ─── Currency / Numeric Parser ────────────────────────────────────────────────
_CURRENCY_RE = re.compile(r"[\$€£¥₦,\s]")

def parse_numeric(val) -> float | None:
    """Strip currency symbols, commas, and whitespace; return float or None."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ("", "N/A", "NA", "n/a", "na", "None", "null", "NULL", "-", "#N/A"):
        return None
    cleaned = _CURRENCY_RE.sub("", s)
    try:
        return float(cleaned)
    except ValueError:
        return None


# ─── Date Parser ─────────────────────────────────────────────────────────────
_DATE_FMTS = [
    "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y",
    "%d-%m-%Y", "%Y/%m/%d",
    "%b %d %Y", "%B %d %Y",
    "%d %b %Y", "%d %B %Y",
    "%Y%m%d",
]

def parse_date(val) -> pd.Timestamp | None:
    if pd.isna(val):
        return None
    s = str(val).strip()
    for fmt in _DATE_FMTS:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            continue
    try:
        return pd.Timestamp(val)
    except Exception:
        return None


# ─── Smart Type Detector ──────────────────────────────────────────────────────
def detect_column_types(df: pd.DataFrame) -> dict:
    types = {}
    for col in df.columns:
        series = df[col].dropna().astype(str).str.strip()
        if series.empty:
            types[col] = "empty"
            continue

        bool_vals = {"true","false","yes","no","1","0","y","n"}
        if series.str.lower().isin(bool_vals).mean() > 0.85:
            types[col] = "boolean"
            continue

        numeric_hits = series.apply(lambda x: parse_numeric(x) is not None).mean()
        if numeric_hits > 0.70:
            types[col] = "numeric"
            continue

        sample = series.sample(min(30, len(series)), random_state=42)
        date_hits = sample.apply(lambda x: parse_date(x) is not None).mean()
        if date_hits > 0.60:
            types[col] = "date"
            continue

        col_lower = col.lower()
        id_keywords = {"id","code","ref","number","no","num","email","phone","mobile","tel"}
        if any(kw in col_lower for kw in id_keywords):
            types[col] = "identifier"
            continue

        nunique = series.nunique()
        if nunique <= max(10, len(series) * 0.15):
            types[col] = "categorical"
        else:
            types[col] = "text"

    return types


# ─── Outlier Detector ──────────────────────────────────────────────────────────
def detect_outliers(series: pd.Series, factor: float = 3.0) -> pd.Series:
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return pd.Series(False, index=series.index)
    lower = q1 - factor * iqr
    upper = q3 + factor * iqr
    return (series < lower) | (series > upper)


# ─── Core Cleaner ────────────────────────────────────────────────────────────
class DataCleanPro:
    SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls", ".xlsm", ".xlsb"}

    def __init__(self, filepath: str, output_dir: str = "output",
                 outlier_factor: float = 3.0, silent: bool = False,
                 sheet_name=0):
        self.filepath       = Path(filepath)
        self.output_dir     = Path(output_dir)
        self.outlier_factor = outlier_factor
        self.silent         = silent
        self.sheet_name     = sheet_name
        self.log            = []
        self.stats          = {}

    def run(self) -> dict:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        stem = self.filepath.stem
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")

        if not self.silent:
            step_banner(f"Loading  →  {self.filepath.name}")

        df_raw = self._load()
        df     = df_raw.copy()
        n_rows_original = len(df)
        n_cols_original = len(df.columns)

        stages = [
            ("Stripping whitespace",           self._strip_whitespace),
            ("Standardising null markers",     self._standardise_nulls),
            ("Removing blank rows/cols",       self._remove_blank),
            ("Detecting column types",         self._detect_and_cast),
            ("Standardising category case",    self._standardise_categories),
            ("Parsing & standardising dates",  self._parse_dates),
            ("Parsing currency / numerics",    self._parse_currencies),
            ("Removing duplicates",            self._remove_duplicates),
            ("Flagging outliers",              self._flag_outliers),
            ("Imputing missing values",        self._impute_missing),
            ("Generating column summary",      self._column_summary),
        ]

        total = len(stages)
        for i, (label, fn) in enumerate(stages, 1):
            if not self.silent:
                progress_bar("Cleaning", total, i - 1)
                time.sleep(0.05)
            try:
                df = fn(df)
            except Exception as exc:
                self._log("error", label, str(exc))
            if not self.silent:
                progress_bar("Cleaning", total, i)

        self.stats = {
            "file":            self.filepath.name,
            "rows_in":         n_rows_original,
            "cols_in":         n_cols_original,
            "rows_out":        len(df),
            "cols_out":        len(df.columns),
            "rows_removed":    n_rows_original - len(df),
            "missing_before":  int(df_raw.isna().sum().sum()),
            "missing_after":   int(df.isna().sum().sum()),
            "processed_at":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        excel_path  = self.output_dir / f"{stem}_cleaned_{ts}.xlsx"
        report_path = self.output_dir / f"{stem}_report_{ts}.html"

        self._save_excel(df, excel_path)
        self._save_report(df, report_path)

        if not self.silent:
            print()
            self._print_summary(excel_path, report_path)

        return {
            "cleaned_df":  df,
            "excel":       str(excel_path),
            "report":      str(report_path),
            "stats":       self.stats,
            "log":         self.log,
        }

    def _load(self) -> pd.DataFrame:
        ext = self.filepath.suffix.lower()

        if ext in (".xlsx", ".xls", ".xlsm", ".xlsb"):
            try:
                df = pd.read_excel(
                    self.filepath,
                    sheet_name=self.sheet_name,
                    engine="openpyxl" if ext != ".xls" else "xlrd",
                )
                self._log("ok", "File loaded",
                          f"{len(df)} rows · {len(df.columns)} columns · format=Excel({ext})")
                return df
            except Exception as exc:
                raise ValueError(f"Cannot read Excel file {self.filepath}: {exc}")

        encodings = ["utf-8", "latin-1", "cp1252", "utf-8-sig"]
        for enc in encodings:
            try:
                df = pd.read_csv(self.filepath, encoding=enc, low_memory=False)
                self._log("ok", "File loaded",
                          f"{len(df)} rows · {len(df.columns)} columns · encoding={enc}")
                return df
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        raise ValueError(f"Cannot decode {self.filepath}. Try saving as UTF-8.")

    def _strip_whitespace(self, df: pd.DataFrame) -> pd.DataFrame:
        df.columns = df.columns.str.strip()
        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace("nan", np.nan)
        self._log("ok", "Whitespace stripped", "Column names + string cells cleaned")
        return df

    def _standardise_nulls(self, df: pd.DataFrame) -> pd.DataFrame:
        null_markers = ["N/A", "NA", "n/a", "na", "None", "none", "null",
                        "NULL", "NaN", "-", "#N/A", "#VALUE!", "?", ".",
                        "", " ", "  "]
        before = int(df.isna().sum().sum())
        df = df.replace(null_markers, np.nan)
        after = int(df.isna().sum().sum())
        self._log("ok", "Null markers standardised",
                  f"{after - before} additional nulls identified")
        return df

    def _remove_blank(self, df: pd.DataFrame) -> pd.DataFrame:
        blank_rows = df.isna().all(axis=1).sum()
        df = df.dropna(how="all")
        blank_cols = df.isna().all(axis=0).sum()
        df = df.dropna(axis=1, how="all")
        self._log("ok", "Blank rows / columns removed",
                  f"{blank_rows} empty rows · {blank_cols} empty columns dropped")
        return df

    def _detect_and_cast(self, df: pd.DataFrame) -> pd.DataFrame:
        self.col_types = detect_column_types(df)
        summary = ", ".join(f"{k}:{v}" for k, v in
                            pd.Series(self.col_types).value_counts().items())
        self._log("ok", "Column types detected", summary)
        return df

    def _standardise_categories(self, df: pd.DataFrame) -> pd.DataFrame:
        cat_cols = [c for c, t in self.col_types.items()
                    if t == "categorical" and c in df.columns]
        changed = 0
        for col in cat_cols:
            original = df[col].copy()
            df[col] = df[col].astype(str).str.strip().str.title()
            df[col] = df[col].replace("Nan", np.nan)
            if not original.equals(df[col]):
                changed += 1
        self._log("ok", "Categories standardised",
                  f"{changed} columns → Title Case applied")
        return df

    def _parse_dates(self, df: pd.DataFrame) -> pd.DataFrame:
        date_cols = [c for c, t in self.col_types.items()
                     if t == "date" and c in df.columns]
        fixed = 0
        for col in date_cols:
            df[col] = df[col].apply(parse_date)
            df[col] = df[col].apply(
                lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) else np.nan
            )
            fixed += 1
        self._log("ok", "Dates standardised",
                  f"{fixed} date columns → YYYY-MM-DD format")
        return df

    def _parse_currencies(self, df: pd.DataFrame) -> pd.DataFrame:
        num_cols = [c for c, t in self.col_types.items()
                    if t == "numeric" and c in df.columns]
        converted = 0
        for col in num_cols:
            new_vals = df[col].apply(parse_numeric)
            if new_vals.notna().sum() > 0:
                df[col] = new_vals
                converted += 1
        self._log("ok", "Currency / numerics parsed",
                  f"{converted} columns stripped of symbols → float")
        return df

    def _remove_duplicates(self, df: pd.DataFrame) -> pd.DataFrame:
        before = len(df)
        df = df.drop_duplicates()
        after = len(df)
        removed = before - after
        if removed:
            self._log("warn", "Duplicate rows removed", f"{removed} exact duplicates dropped")
        else:
            self._log("ok", "Duplicate check", "No duplicate rows found")
        return df

    def _flag_outliers(self, df: pd.DataFrame) -> pd.DataFrame:
        num_cols = [c for c, t in self.col_types.items()
                    if t == "numeric" and c in df.columns]
        outlier_report = []
        for col in num_cols:
            series = pd.to_numeric(df[col], errors="coerce").dropna()
            if len(series) < 4:
                continue
            mask = detect_outliers(series, self.outlier_factor)
            n_out = int(mask.sum())
            if n_out > 0:
                outlier_report.append(f"{col}: {n_out} outlier(s)")
                flag_col = f"_outlier_{col}"
                df[flag_col] = False
                df.loc[mask.index[mask], flag_col] = True

        if outlier_report:
            self._log("warn", "Outliers flagged",
                      "; ".join(outlier_report) + " — flagged in _outlier_* columns")
        else:
            self._log("ok", "Outlier check", "No statistical outliers detected")
        return df

    def _impute_missing(self, df: pd.DataFrame) -> pd.DataFrame:
        report = []
        for col in df.columns:
            if col.startswith("_outlier_"):
                continue
            n_missing = int(df[col].isna().sum())
            if n_missing == 0:
                continue
            col_type = self.col_types.get(col, "text")
            if col_type == "numeric":
                median_val = pd.to_numeric(df[col], errors="coerce").median()
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(median_val)
                report.append(f"{col}(median={median_val:.2f})")
            elif col_type == "categorical":
                mode_val = df[col].mode()
                if not mode_val.empty:
                    df[col] = df[col].fillna(mode_val.iloc[0])
                    report.append(f"{col}(mode='{mode_val.iloc[0]}')")
        if report:
            self._log("ok", "Missing values imputed", "; ".join(report))
        else:
            self._log("ok", "Missing value check",
                      "No imputation needed (or columns are non-numeric)")
        return df

    def _column_summary(self, df: pd.DataFrame) -> pd.DataFrame:
        self.column_stats = []
        for col in df.columns:
            if col.startswith("_outlier_"):
                continue
            col_data = df[col]
            inferred = self.col_types.get(col, "?")
            missing_n = int(col_data.isna().sum())
            missing_pct = round(100 * missing_n / max(len(df), 1), 1)
            unique_n = int(col_data.nunique())
            if inferred == "numeric":
                numeric_s = pd.to_numeric(col_data, errors="coerce")
                sample = f"min={numeric_s.min():.2f} · max={numeric_s.max():.2f} · mean={numeric_s.mean():.2f}"
            elif inferred == "categorical":
                top = col_data.value_counts().head(3).index.tolist()
                sample = "Top: " + ", ".join(str(v) for v in top)
            else:
                sample = str(col_data.dropna().iloc[0]) if col_data.dropna().any() else "—"
            self.column_stats.append({
                "Column": col,
                "Type": inferred,
                "Missing": f"{missing_n} ({missing_pct}%)",
                "Unique": unique_n,
                "Sample / Stats": sample[:80],
            })
        self._log("ok", "Column summary generated", f"{len(self.column_stats)} columns analysed")
        return df

    def _save_excel(self, df: pd.DataFrame, path: Path) -> None:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Cleaned Data", index=False)
            log_df = pd.DataFrame(self.log)
            log_df.to_excel(writer, sheet_name="Cleaning Log", index=False)
            if hasattr(self, "column_stats"):
                pd.DataFrame(self.column_stats).to_excel(
                    writer, sheet_name="Column Summary", index=False)
            stats_df = pd.DataFrame([self.stats]).T.reset_index()
            stats_df.columns = ["Metric", "Value"]
            stats_df.to_excel(writer, sheet_name="Run Stats", index=False)
        self._log("ok", "Excel saved", str(path))

    def _save_report(self, df: pd.DataFrame, path: Path) -> None:
        html = self._build_html_report(df)
        path.write_text(html, encoding="utf-8")
        self._log("ok", "HTML report saved", str(path))

    def _print_summary(self, excel: Path, report: Path) -> None:
        s = self.stats
        print(f"\n{C['header']}{'═'*62}")
        print(f"  CLEANING COMPLETE — {self.filepath.name}")
        print(f"{'═'*62}{C['reset']}")
        table = [
            ["Rows processed",    s['rows_in']],
            ["Rows output",       s['rows_out']],
            ["Rows removed",      s['rows_removed']],
            ["Missing (before)",  s['missing_before']],
            ["Missing (after)",   s['missing_after']],
            ["Missing fixed",     s['missing_before'] - s['missing_after']],
        ]
        print(tabulate(table, headers=["Metric", "Value"],
                       tablefmt="rounded_outline"))
        print(f"\n  {C['ok']}📊 Excel  →{C['reset']} {excel}")
        print(f"  {C['ok']}📋 Report →{C['reset']} {report}")
        print(f"\n  {C['dim']}{BRAND}{C['reset']}\n")

    def _build_html_report(self, df: pd.DataFrame) -> str:
        s = self.stats
        now = s['processed_at']

        log_rows = ""
        for entry in self.log:
            icon = {"ok":"✓","warn":"⚠","error":"✗"}.get(entry["level"], "·")
            colour = {"ok":"#22c55e","warn":"#f59e0b","error":"#ef4444"}.get(entry["level"],"#888")
            log_rows += (
                f'<tr><td style="color:{colour};font-weight:600">{icon} {entry["stage"]}</td>'
                f'<td>{entry["detail"]}</td></tr>'
            )

        col_rows = ""
        if hasattr(self, "column_stats"):
            for row in self.column_stats:
                col_rows += (
                    f'<tr><td style="font-family:monospace">{row["Column"]}</td>'
                    f'<td><span style="background:#e0f2fe;padding:2px 8px;border-radius:20px;font-size:12px">{row["Type"]}</span></td>'
                    f'<td>{row["Missing"]}</td>'
                    f'<td>{row["Unique"]}</td>'
                    f'<td style="color:#64748b;font-size:13px">{row["Sample / Stats"]}</td>'
                    f'</tr>'
                )

        # FIXED SECTION - proper indentation
        preview_cols = [c for c in df.columns if not c.startswith("_outlier_")]
        preview_df = df[preview_cols].head(10)
        th = "".join(f"<th>{c}</th>" for c in preview_df.columns)
        
        td_rows = ""
        for _, row in preview_df.iterrows():
            tds_list = []
            for v in row:
                if pd.notna(v):
                    tds_list.append(f'<td>{str(v)}</td>')
                else:
                    tds_list.append('<td><em style="color:#cbd5e1">null</em></td>')
            tds = "".join(tds_list)
            td_rows += f'<tr>{tds}</tr>'

        efficiency = round(100 * (1 - s['missing_after'] / max(s['missing_before'], 1)), 1)

        return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DataClean Pro Report — {s['file']}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:'Segoe UI',system-ui,sans-serif;background:#f8fafc;color:#1e293b;line-height:1.6}}
  a{{color:#2563eb;text-decoration:none}}
  .header{{background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);color:#fff;padding:2.5rem 3rem}}
  .header h1{{font-size:2rem;font-weight:800}}
  .brand-link{{color:#60a5fa;font-weight:600}}
  .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:1rem;padding:2rem 3rem 0}}
  .card{{background:#fff;border-radius:12px;padding:1.25rem 1.5rem;border:1px solid #e2e8f0}}
  .card .num{{font-size:2rem;font-weight:800;color:#0f172a}}
  .card .lbl{{font-size:0.8rem;color:#64748b}}
  .card.green .num{{color:#16a34a}}
  .section{{margin:2rem 3rem}}
  .section-title{{font-size:1.1rem;font-weight:700;margin-bottom:1rem;border-bottom:2px solid #e2e8f0;padding-bottom:.5rem}}
  table{{width:100%;border-collapse:collapse;font-size:0.875rem;background:#fff;border:1px solid #e2e8f0}}
  th{{background:#f1f5f9;padding:.75rem 1rem;text-align:left}}
  td{{padding:.65rem 1rem;border-top:1px solid #f1f5f9}}
  .footer{{background:#0f172a;color:#64748b;padding:2rem 3rem;margin-top:3rem}}
  .footer a{{color:#60a5fa}}
</style>
</head>
<body>

<div class="header">
  <h1>📊 DataClean Pro — Audit Report</h1>
  <div class="sub">
    File: <strong>{s['file']}</strong> &nbsp;·&nbsp; Processed: {now}<br>
    Powered by <a class="brand-link" href="https://eduxellence.org" target="_blank">Eduxellence Analytics</a>
  </div>
</div>

<div class="cards">
  <div class="card"><div class="num">{s['rows_in']}</div><div class="lbl">Rows in</div></div>
  <div class="card green"><div class="num">{s['rows_out']}</div><div class="lbl">Rows out</div></div>
  <div class="card"><div class="num">{s['rows_removed']}</div><div class="lbl">Rows removed</div></div>
  <div class="card"><div class="num">{s['missing_before']}</div><div class="lbl">Missing (before)</div></div>
  <div class="card green"><div class="num">{s['missing_after']}</div><div class="lbl">Missing (after)</div></div>
  <div class="card green"><div class="num">{efficiency}%</div><div class="lbl">Missing fixed</div></div>
</div>

<div class="section">
  <div class="section-title">🔧 Cleaning Log</div>
  <table><thead><tr><th>Stage</th><th>Detail</th></tr></thead><tbody>{log_rows}</tbody></table>
</div>

<div class="section">
  <div class="section-title">📋 Column Intelligence Report</div>
  <div class="scroll-x"><table><thead><tr><th>Column</th><th>Type</th><th>Missing</th><th>Unique</th><th>Sample / Stats</th></tr></thead><tbody>{col_rows}</tbody></table></div>
</div>

<div class="section">
  <div class="section-title">👁 Cleaned Data Preview (first 10 rows)</div>
  <div class="scroll-x"><table><thead><tr>{th}</tr></thead><tbody>{td_rows}</tbody></table></div>
</div>

<div class="footer">
  <p>Generated by <strong>DataClean Pro v1.0</strong> — a free tool by <a href="https://eduxellence.org" target="_blank">Eduxellence Analytics</a></p>
  <p>Need advanced analytics? <a href="https://eduxellence.org/#contact" target="_blank">Book a free 15-minute Data Strategy Audit</a></p>
</div>

</body>
</html>"""

    def _log(self, level: str, stage: str, detail: str) -> None:
        self.log.append({"level": level, "stage": stage, "detail": detail})


# ════════════════════════════════════════════════════════════════════
# GUI (Tkinter) - requires tkinterdnd2 for drag-and-drop
# ════════════════════════════════════════════════════════════════════
def launch_gui() -> None:
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, scrolledtext
    except ImportError:
        print("Tkinter not available. On Linux: sudo apt-get install python3-tk")
        sys.exit(1)

    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        _DND_AVAILABLE = True
    except ImportError:
        _DND_AVAILABLE = False

    RootClass = TkinterDnD.Tk if _DND_AVAILABLE else tk.Tk
    root = RootClass()
    root.title("DataClean Pro — Eduxellence Analytics")
    root.geometry("780x620")
    root.configure(bg="#f8fafc")

    NAVY, BLUE, GREEN, WHITE = "#0f172a", "#2563eb", "#16a34a", "#ffffff"

    header = tk.Frame(root, bg=NAVY, height=70)
    header.pack(fill="x")
    tk.Label(header, text="📊 DataClean Pro", font=("Segoe UI", 18, "bold"),
             bg=NAVY, fg=WHITE).pack(side="left", padx=20, pady=12)
    tk.Label(header, text="by Eduxellence Analytics", font=("Segoe UI", 9),
             bg=NAVY, fg="#94a3b8").pack(side="left")

    body = tk.Frame(root, bg="#f8fafc")
    body.pack(fill="both", expand=True, padx=20, pady=16)

    file_var, output_var, factor_var, sheet_var = tk.StringVar(), tk.StringVar(value="output"), tk.StringVar(value="3.0"), tk.StringVar(value="0")

    drop_frame = tk.Frame(body, bg=WHITE, relief="groove", bd=2)
    drop_frame.pack(fill="x", pady=(0,12))
    drop_label = tk.Label(drop_frame,
        text="Drag & drop CSV/Excel file here" if _DND_AVAILABLE else "Select a file using Browse",
        bg=WHITE, fg="#64748b", pady=22)
    drop_label.pack()

    def make_row(parent, label_text, widget_factory):
        f = tk.Frame(parent, bg="#f8fafc")
        f.pack(fill="x", pady=4)
        tk.Label(f, text=label_text, font=("Segoe UI", 10, "bold"), bg="#f8fafc", width=16, anchor="w").pack(side="left")
        widget_factory(f)

    def file_picker(parent):
        tk.Entry(parent, textvariable=file_var, width=52).pack(side="left", padx=(0,6))
        def browse():
            p = filedialog.askopenfilename(filetypes=[("Data files", "*.csv *.xlsx *.xls")])
            if p:
                file_var.set(p)
                drop_label.config(text=f"✓ {Path(p).name}", fg=GREEN)
        tk.Button(parent, text="Browse", command=browse).pack(side="left")

    def out_picker(parent):
        tk.Entry(parent, textvariable=output_var, width=52).pack(side="left", padx=(0,6))
        def browse():
            p = filedialog.askdirectory()
            if p:
                output_var.set(p)
        tk.Button(parent, text="Browse", command=browse).pack(side="left")

    make_row(body, "Input file:", file_picker)
    make_row(body, "Output folder:", out_picker)
    make_row(body, "Outlier factor:", lambda p: tk.Entry(p, textvariable=factor_var, width=8).pack(side="left"))
    make_row(body, "Excel sheet:", lambda p: tk.Entry(p, textvariable=sheet_var, width=16).pack(side="left"))

    tk.Label(body, text="Cleaning log:", font=("Segoe UI", 10, "bold"), bg="#f8fafc", anchor="w").pack(fill="x", pady=(8,2))
    log_box = scrolledtext.ScrolledText(body, height=10, font=("Courier New", 9), bg="#0f172a", fg="#e2e8f0")
    log_box.pack(fill="both", expand=True)

    prog_var = tk.DoubleVar()
    ttk.Progressbar(body, variable=prog_var, maximum=100).pack(fill="x", pady=(8,0))

    def run_clean():
        if not file_var.get():
            messagebox.showerror("Error", "Select a file")
            return
        log_box.delete("1.0", "end")
        prog_var.set(0)

        def task():
            cleaner = DataCleanPro(file_var.get(), output_dir=output_var.get(),
                                   outlier_factor=float(factor_var.get()), silent=True)
            stage_count = [0]
            def gui_log(level, stage, detail):
                cleaner._log(level, stage, detail)
                log_box.insert("end", f"  {stage}: {detail}\n")
                log_box.see("end")
                stage_count[0] += 1
                prog_var.set(stage_count[0] / 11 * 100)
            cleaner._log = gui_log
            try:
                result = cleaner.run()
                messagebox.showinfo("Complete", f"Done!\nExcel: {result['excel']}\nReport: {result['report']}")
            except Exception as e:
                messagebox.showerror("Error", str(e))
        import threading
        threading.Thread(target=task, daemon=True).start()

    tk.Button(root, text="▶ Clean Data", font=("Segoe UI", 12, "bold"), bg=GREEN, fg=WHITE,
              command=run_clean, padx=20, pady=10).pack(pady=12)
    root.mainloop()


# ════════════════════════════════════════════════════════════════════
# Google Colab helper
# ════════════════════════════════════════════════════════════════════
COLAB_CELL = '''
# DataClean Pro — Google Colab Edition
import subprocess, sys, urllib.request, os
subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "colorama", "tabulate", "openpyxl"])
urllib.request.urlretrieve("https://raw.githubusercontent.com/eduxellence/dataclean-pro/main/dataclean_pro.py", "dataclean_pro.py")
from google.colab import files
print("Upload your CSV/Excel file:")
uploaded = files.upload()
filename = list(uploaded.keys())[0]
from dataclean_pro import DataCleanPro
cleaner = DataCleanPro(filename, output_dir="output")
result = cleaner.run()
files.download(result["excel"])
files.download(result["report"])
print("Done!")
'''

def print_colab_cell():
    print(COLAB_CELL)

def save_colab_notebook(output_dir="."):
    nb = {"nbformat": 4, "nbformat_minor": 5, "metadata": {}, "cells": [{"cell_type": "code", "metadata": {}, "execution_count": None, "source": COLAB_CELL.strip().splitlines()}]}
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    Path(output_dir, "dataclean_pro_colab.ipynb").write_text(json.dumps(nb))


# ════════════════════════════════════════════════════════════════════
# Batch runner
# ════════════════════════════════════════════════════════════════════
def run_batch(folder: str, output_dir: str, silent: bool) -> None:
    folder = Path(folder)
    all_files = list(folder.rglob("*.csv")) + list(folder.rglob("*.xlsx"))
    if not all_files:
        err(f"No CSV/Excel files in {folder}")
        return
    print(f"\nBatch mode: {len(all_files)} files\n")
    results = []
    for i, f in enumerate(all_files, 1):
        print(f"[{i}/{len(all_files)}] {f.name}")
        cleaner = DataCleanPro(f, output_dir=output_dir, silent=silent)
        r = cleaner.run()
        results.append({"file": f.name, **r["stats"]})
    print(tabulate(results, headers="keys", tablefmt="rounded_outline"))


# ─── Interactive mode ──────────────────────────────────────────────────
def interactive_mode(output_dir: str) -> None:
    print(BANNER)
    all_files = list(Path(".").glob("*.csv")) + list(Path(".").glob("*.xlsx"))
    if all_files:
        for i, p in enumerate(all_files[:20], 1):
            print(f"  [{i}] {p.name}")
    filepath = input("Enter file path or number: ").strip()
    if filepath.isdigit() and 1 <= int(filepath) <= len(all_files):
        filepath = str(all_files[int(filepath)-1])
    cleaner = DataCleanPro(filepath, output_dir=output_dir)
    cleaner.run()


# ─── CLI entry point ─────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(prog="dataclean_pro", description="DataClean Pro — automated CSV/Excel cleaning")
    parser.add_argument("file", nargs="?")
    parser.add_argument("--output", "-o", default="output")
    parser.add_argument("--batch", "-b", metavar="FOLDER")
    parser.add_argument("--silent", "-s", action="store_true")
    parser.add_argument("--gui", "-g", action="store_true")
    parser.add_argument("--colab", "-c", action="store_true")
    args = parser.parse_args()

    if args.colab:
        print_colab_cell()
        save_colab_notebook(args.output)
        return
    if args.gui:
        launch_gui()
        return
    if not args.silent:
        print(BANNER)
    if args.batch:
        run_batch(args.batch, args.output, args.silent)
    elif args.file:
        cleaner = DataCleanPro(args.file, output_dir=args.output, silent=args.silent)
        cleaner.run()
    else:
        interactive_mode(args.output)


if __name__ == "__main__":
    main()