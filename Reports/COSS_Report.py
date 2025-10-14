# -*- coding: utf-8 -*-
from pathlib import Path
import datetime as dt
import pandas as pd

# ================== PATHS ==================
ONE_DRIVE = Path.home() / "OneDrive - Reconnect Community Health Services (1)"
CLEAN_DIR = ONE_DRIVE / "data" / "Clean Data" / "Treat"
PROGFIN_STEM = "rpt_ProgStaffFinance"   # use newest .csv/.xlsx/.xls

# ================== CONFIG ==================
TARGET_PROGRAM = "COSS Follow Up"
DATE_COL = "Interaction Start Date"
PROGRAM_COL = "Program"

# Only add Q columns when all 3 months exist (True) vs. sum whatever exists so far (False)
STRICT_FULL_QUARTER = True


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
        return dt.date(today.year, 4, 1), dt.date(today.year + 1, 3, 31)
    return dt.date(today.year - 1, 4, 1), dt.date(today.year, 3, 31)


def _month_range(year: int, month: int) -> tuple[dt.date, dt.date]:
    first = dt.date(year, month, 1)
    ny, nm = year + (month // 12), (month % 12) + 1
    last = dt.date(ny, nm, 1) - dt.timedelta(days=1)
    return first, last


def _fmt_col(d: dt.date) -> str:
    # "YYYY-MM"
    return f"{d.year:04d}-{d.month:02d}"


def make_visits_monthly_with_quarters() -> pd.DataFrame:
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

    # 4) FY window and cap to end of current month
    today = dt.date.today()
    fy_start, fy_end = _current_fy_bounds(today)
    df = df[(df[DATE_COL] >= fy_start) & (df[DATE_COL] <= fy_end)].copy()

    cap_to = (dt.date(today.year + (today.month // 12), (today.month % 12) + 1, 1)
              - dt.timedelta(days=1))

    # 5) Build monthly list Apr..Mar (only months up to cap_to)
    months = []
    y, m = fy_start.year, fy_start.month  # Apr
    for _ in range(12):
        first, last = _month_range(y, m)
        if last > cap_to:
            break
        months.append((first, last))
        y, m = (y + (m // 12), (m % 12) + 1)

    month_labels = [_fmt_col(last) for (first, last) in months]

    # 6) Compute MONTHLY (non-cumulative) visit counts
    visits_monthly = []
    for first, last in months:
        mdf = df[(df[DATE_COL] >= first) & (df[DATE_COL] <= last)]
        visits_monthly.append(len(mdf))  # each row = a visit record

    # 7) Base output with monthly values
    out = pd.DataFrame({"Metric": ["Visits (Monthly)"]})
    for i, lbl in enumerate(month_labels):
        out[lbl] = [visits_monthly[i]]

    # 8) Add fiscal quarter sums (Apr–Jun, Jul–Sep, Oct–Dec, Jan–Mar)
    # FY month numbers: Apr=1, May=2, ..., Mar=12
    fy_idx_by_label = {}
    for idx, (first, last) in enumerate(months, start=1):
        fy_idx_by_label[idx] = month_labels[idx - 1]

    def add_quarter(qname: str, fy_indices: list[int]):
        if not all(i in fy_idx_by_label for i in fy_indices):
            if STRICT_FULL_QUARTER:
                return None
        existing = [fy_idx_by_label[i] for i in fy_indices if i in fy_idx_by_label]
        if not existing:
            return None
        q_total = sum(out[label].iloc[0] for label in existing)
        out[qname] = [q_total]
        return existing[-1]  # place Q right after the last month of the quarter

    quarter_specs = [("Q1", [1, 2, 3]), ("Q2", [4, 5, 6]), ("Q3", [7, 8, 9]), ("Q4", [10, 11, 12])]
    final_cols = ["Metric"] + month_labels
    for qname, idxs in quarter_specs:
        after_lbl = add_quarter(qname, idxs)
        if after_lbl:
            pos = final_cols.index(after_lbl)
            final_cols = final_cols[:pos + 1] + [qname] + final_cols[pos + 1:]

    out = out[final_cols]
    return out


def run(save_csv: str | None = None, verbose: bool = True) -> pd.DataFrame:
    """
    Build a MONTHLY visits table and insert fiscal Q1–Q4 columns
    as sums of Apr–Jun, Jul–Sep, Oct–Dec, Jan–Mar.
    """
    table = make_visits_monthly_with_quarters()
    if verbose:
        print(table.to_string(index=False))
    if save_csv:
        Path(save_csv).parent.mkdir(parents=True, exist_ok=True)
        table.to_csv(save_csv, index=False, encoding="utf-8-sig")
    return table


if __name__ == "__main__":
    run()  # or: run(save_csv=r"C:\path\to\COSS_FollowUp_Visits_Monthly_with_Q.csv")
