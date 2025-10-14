# Reports/VHA_Monthly_Metrics.py
# Computes:
#   - total_yes_no
#   - total_yes
#   - Individual_served
#   - count_first_in_prev_month
#   - num_of_discharges
#
# run() returns a 2-col DataFrame (Metric | Value) so it renders in your Output table.
# main() prints a friendly summary for CLI runs.

import re
import pandas as pd
from pathlib import Path
from typing import Dict, Tuple
import os

# ================== CONFIG ==================
user = os.environ.get("UserName") or os.environ.get("USERNAME") or os.getlogin()
FOLDER = Path(r"C:\Users\=\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\AlayaCare".replace("=", user))

# ================== HELPERS ==================
def _find_csv(base_name: str) -> Path:
    exact = FOLDER / f"{base_name}.csv"
    if exact.exists():
        return exact
    base_lower = base_name.lower()
    for p in sorted(FOLDER.glob("*.csv")):
        if p.stem.lower().startswith(base_lower):
            return p
    raise FileNotFoundError(f"Couldn't find CSV for '{base_name}' in {FOLDER}")

def _read_csv_any(path: Path) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(
                path, dtype=str, encoding=enc, engine="python",
                on_bad_lines="skip", keep_default_na=False
            )
        except Exception:
            pass
    # last fallback
    return pd.read_csv(path, dtype=str, engine="python",
                       on_bad_lines="skip", keep_default_na=False)

def _find_col(df: pd.DataFrame, aliases: list[str]) -> str:
    """Find a column by aliases, ignoring spaces/underscore/case."""
    def norm(s: str) -> str:
        return re.sub(r"[\W_]+", "", (s or "").lower())
    norm_map = {norm(c): c for c in df.columns}
    for a in aliases:
        na = norm(a)
        if na in norm_map:
            return norm_map[na]
    raise KeyError(f"None of {aliases} found. Columns present: {list(df.columns)}")

def _parse_date_col(df: pd.DataFrame, col: str) -> pd.Series:
    # Parse to UTC then drop tz (naive) for consistent comparisons
    return pd.to_datetime(df[col], errors="coerce", utc=True).dt.tz_convert(None)

def _prev_month_window(today: pd.Timestamp | None = None) -> Tuple[pd.Timestamp, pd.Timestamp, pd.Timestamp]:
    """Return (prev_start, prev_end, month_start_of_current)."""
    if today is None:
        today = pd.Timestamp.today()
    month_start = pd.Timestamp(today.year, today.month, 1)
    prev_end = month_start - pd.Timedelta(days=1)
    prev_start = pd.Timestamp(prev_end.year, prev_end.month, 1)
    return prev_start, prev_end, month_start

# ================== CORE ==================
def _compute_metrics() -> Dict[str, int]:
    visits_df = _read_csv_any(_find_csv("Visits"))

    # Resolve key columns
    funder_col  = _find_col(visits_df, ["FunderCode", "Funder Code", "Funder"])
    date_col    = _find_col(visits_df, ["Date", "Visit Date"])
    status_col  = _find_col(visits_df, ["Visit Status", "Status"])
    visitid_col = _find_col(visits_df, ["Visit Visit ID", "Visit ID", "VisitID"])
    client_col  = _find_col(visits_df, ["AlaycareID", "AlayaCareID", "ClientID", "Client Id"])

    # Filter VHA
    visits_df[funder_col] = visits_df[funder_col].astype(str).str.strip().str.upper()
    vha_all = visits_df[visits_df[funder_col] == "VHA"].copy()

    # Parse dates once
    vha_all["__Date"] = _parse_date_col(vha_all, date_col)

    # Time windows
    prev_start, prev_end, month_start = _prev_month_window()
    prev2_end   = prev_start - pd.Timedelta(days=1)
    prev2_start = pd.Timestamp(prev2_end.year, prev2_end.month, 1)

    # Windowed slices
    mask_prev  = (vha_all["__Date"] >= prev_start) & (vha_all["__Date"] <= prev_end)      # M1: previous full month
    mask_prev2 = (vha_all["__Date"] >= prev2_start) & (vha_all["__Date"] <= prev2_end)    # M2: month before previous

    vha_prev  = vha_all.loc[mask_prev].copy()
    vha_prev2 = vha_all.loc[mask_prev2].copy()

    # total_yes_no & total_yes (distinct Visit IDs with YES/NO in M1)
    vha_prev[status_col] = vha_prev[status_col].astype(str).str.strip().str.upper()
    pm_yn = vha_prev[vha_prev[status_col].isin(["YES", "NO"])]

    total_yes_no = pm_yn[visitid_col].nunique()
    total_yes    = pm_yn.loc[pm_yn[status_col] == "YES", visitid_col].nunique()

    # Individual_served: distinct clients in M1
    Individual_served = vha_prev[client_col].dropna().astype(str).nunique()

    # count_first_in_prev_month: first-ever VHA visit landed in M1
    first_visits = (
        vha_all.dropna(subset=[client_col, "__Date"])
               .groupby(client_col, dropna=False)["__Date"]
               .min()
               .rename("FirstVisitDate")
               .reset_index()
    )
    in_prev_full_month = (first_visits["FirstVisitDate"] >= prev_start) & (first_visits["FirstVisitDate"] < month_start)
    count_first_in_prev_month = int(in_prev_full_month.sum())

    # num_of_discharges: drop-offs from M2 to M1 (only if M1 < M2)
    served_prev      = set(vha_prev[client_col].dropna().astype(str).unique())   # M1
    served_prev_prev = set(vha_prev2[client_col].dropna().astype(str).unique())  # M2

    if len(served_prev) < len(served_prev_prev):
        dropped_ids = served_prev_prev - served_prev
        num_of_discharges = len(dropped_ids)
    else:
        num_of_discharges = 0

    return {
        "Visit Include Only Yes": total_yes,
        "Visit Include Both Yes and No": total_yes_no,
        "Individual Served": Individual_served,
        "Number of New Admission": count_first_in_prev_month,
        "Number of Discharges": num_of_discharges,
    }

def run() -> pd.DataFrame:
    """
    Entry for your app. Returns a 2-col DataFrame (Metric | Value)
    so it renders nicely in the Output tab.
    """
    metrics = _compute_metrics()
    rows = [{"Metric": k, "Value": v} for k, v in metrics.items()]
    return pd.DataFrame(rows, columns=["Metric", "Value"])

def main():
    df = run()
    # Pretty CLI printout
    print(df.to_string(index=False))

if __name__ == "__main__":
    main()
