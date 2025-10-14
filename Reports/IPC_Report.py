# -*- coding: utf-8 -*-
import os
from pathlib import Path
import datetime as dt
import pandas as pd

# ================== PATHS ==================
ONE_DRIVE = Path.home() / "OneDrive - Reconnect Community Health Services (1)"
CLEAN_DIR = ONE_DRIVE / "data" / "Clean Data" / "Treat"
PROGFIN_STEM = "rpt_ProgStaffFinance"   # use newest .csv/.xlsx/.xls

# ================== CONFIG ==================
TARGET_PROGRAM = "Primary Care InterProfessional"
DATE_COL = "Interaction Start Date"
PROGRAM_COL = "Program"

# Candidate ID columns for Individuals Served (unique people)
POSSIBLE_ID_COLS = ["ID", "Client ID", "Client Name"]

def _find_latest(path: Path, stem: str) -> Path:
    cands = []
    for ext in (".csv", ".xlsx", ".xls"):
        cands += list(path.glob(f"{stem}*{ext}"))
    if not cands:
        raise FileNotFoundError(f"No '{stem}(.csv/.xlsx/.xls)' in {path}")
    return max(cands, key=lambda p: p.stat().st_mtime)

def _read_any(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    return pd.read_excel(path, dtype=str)

def _current_fy_bounds(today: dt.date) -> tuple[dt.date, dt.date]:
    # FY: Apr 1 ~ Mar 31
    if today.month >= 4:
        return dt.date(today.year, 4, 1), dt.date(today.year+1, 3, 31)
    return dt.date(today.year-1, 4, 1), dt.date(today.year, 3, 31)

def _month_ends_in_fy(fy_start: dt.date, fy_end: dt.date, cap_to: dt.date) -> list[dt.date]:
    # all month ends in FY, but not beyond 'cap_to'
    res = []
    y, m = fy_start.year, fy_start.month
    while True:
        y2, m2 = y + (m // 12), (m % 12) + 1   # next month 1st
        month_end = dt.date(y2, m2, 1) - dt.timedelta(days=1)
        if month_end > fy_end or month_end > cap_to:
            break
        res.append(month_end)
        y, m = y2, m2
    return res

def _fmt_col(d: dt.date) -> str:
    # "YYYY-MM"
    return f"{d.year:04d}-{d.month:02d}"

def make_cumulative_table() -> pd.DataFrame:
    # 1) Load newest rpt_ProgStaffFinance
    p = _find_latest(CLEAN_DIR, PROGFIN_STEM)
    df = _read_any(p)

    # 2) normalize columns
    df = df.rename(columns={c: c.strip() for c in df.columns})
    if DATE_COL not in df or PROGRAM_COL not in df:
        raise KeyError(f"Need columns '{DATE_COL}' and '{PROGRAM_COL}' in {p.name}")

    # 3) parse date & filter program
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce").dt.date
    df = df.dropna(subset=[DATE_COL])

    prog_norm = df[PROGRAM_COL].astype(str).str.strip().str.casefold()
    df = df[prog_norm == TARGET_PROGRAM.strip().casefold()].copy()

    # 4) current FY filter
    today = dt.date.today()
    fy_start, fy_end = _current_fy_bounds(today)
    df = df[(df[DATE_COL] >= fy_start) & (df[DATE_COL] <= fy_end)].copy()

    # 5) choose ID column for Individuals Served
    id_col = next((c for c in POSSIBLE_ID_COLS if c in df.columns), None)
    if not id_col:
        raise KeyError(f"No ID column found. Tried: {POSSIBLE_ID_COLS}")

    # 6) month ends to accumulate through (cap to end of current month)
    cap_to = (dt.date(today.year + (today.month // 12), (today.month % 12) + 1, 1)
              - dt.timedelta(days=1))
    month_ends = _month_ends_in_fy(fy_start, fy_end, cap_to)

    # 7) cumulative metrics
    indiv_cumu, visits_cumu = [], []
    seen_clients = set()

    for mend in month_ends:
        df_up = df[df[DATE_COL] <= mend]

        # Visits = cumulative rows
        visits_cumu.append(len(df_up))

        # Individuals Served = cumulative unique people
        seen_clients.update(df_up[id_col].astype(str).unique())
        indiv_cumu.append(len(seen_clients))

    # 8) build output (metrics as rows; months as columns)
    cols = [_fmt_col(m) for m in month_ends]
    out = pd.DataFrame({
        "Metric": [
            "Individual Served (Cumulative)",
            "Visits (Cumulative)"
        ]
    })
    data_rows = [indiv_cumu, visits_cumu]
    for col_idx, col_name in enumerate(cols):
        out[col_name] = [row[col_idx] for row in data_rows]
    return out

def run(save_csv: str | None = None, verbose: bool = True) -> pd.DataFrame:
    """
    Build the cumulative table and optionally save to CSV.
    Parameters
    ----------
    save_csv : str | None
        If provided, path to save the resulting table as CSV.
    verbose : bool
        If True, print the table.
    """
    table = make_cumulative_table()
    if verbose:
        print(table.to_string(index=False))
    if save_csv:
        Path(save_csv).parent.mkdir(parents=True, exist_ok=True)
        table.to_csv(save_csv, index=False, encoding="utf-8-sig")
    return table

if __name__ == "__main__":
    run()  # or: run(save_csv=r"C:\path\to\PrimaryCare_Cumulative.csv")
