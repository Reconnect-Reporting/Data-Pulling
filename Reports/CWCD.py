# Reports/CWCD.py

import getpass
from pathlib import Path
import pandas as pd

# ============================== Config & Constants ==============================

USER_NAME = getpass.getuser()
DATA_DIR  = Path(fr"C:\Users\{USER_NAME}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat")

MONTH_COLS = ["Apr","May","June","July","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]

INTERVALS = [
    "Service Provider Interactions: 5min - 30min",
    "Service Provider Interactions: 31min - 1hr",
    "Service Provider Interactions: 1hr - 2hr",
    "Service Provider Interactions: 2hr - 5hr",
    "Service Provider Interactions: 5hr - more",
]

# Fiscal-year month → number (Apr starts the FY)
MONTH_NUM = {
    "Apr":4, "May":5, "June":6, "July":7, "Aug":8, "Sep":9, "Oct":10, "Nov":11, "Dec":12,
    "Jan":1, "Feb":2, "Mar":3
}

# ============================== Helpers ==============================

def _pick_latest_census_file(folder: Path) -> Path:
    files = sorted(folder.glob("rpt_Census*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError("No rpt_Census*.csv found")
    return files[0]

def _read_csv_utf8sig(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")

def _fiscal_year_start_end(today: pd.Timestamp) -> tuple[pd.Timestamp, pd.Timestamp, pd.Timestamp]:
    """
    Returns:
      - fy_start (Apr 1 of active FY)
      - eom_last_month (last day of last month)
      - first_of_last_month (1st day of last month)
    """
    today = today.normalize()
    fy_apr1_this_year = pd.Timestamp(year=today.year, month=4, day=1)
    fy_start = fy_apr1_this_year if today >= fy_apr1_this_year else pd.Timestamp(year=today.year - 1, month=4, day=1)
    first_of_this_month = today.replace(day=1)
    eom_last_month = first_of_this_month - pd.Timedelta(days=1)
    first_of_last_month = first_of_this_month - pd.DateOffset(months=1)
    return fy_start, eom_last_month, first_of_last_month

def _build_cumulative_table(visits_df: pd.DataFrame, fy_start: pd.Timestamp, eom_last_month: pd.Timestamp) -> pd.DataFrame:
    """
    Build cumulative (Apr → each month end) table with:
      Through Month | Cumulative Start | Cumulative End | Unique Clients Served (Cum) | Visits (Cum)
    """
    month_ends = pd.date_range(start=fy_start, end=eom_last_month, freq="M")
    rows = []
    for me in month_ends:
        mask_cum = (visits_df["VisitDate"] >= fy_start) & (visits_df["VisitDate"] <= me)
        frame_cum = visits_df.loc[mask_cum]
        rows.append({
            "Through Month": me.strftime("%Y-%m"),
            "Cumulative Start": fy_start.date(),
            "Cumulative End": me.date(),
            "Unique Clients Served (Cumulative)": frame_cum["ID"].nunique(),
            "Visits (Cumulative)": len(frame_cum),
        })
    return pd.DataFrame(rows)

def _mis_stats_cwcd_monthly_totals(mis_stats: pd.DataFrame) -> pd.Series:
    """
    Filter to Program == 'CWCD' and target INTERVALS, keep only month columns,
    and return a Series of month totals (Apr..Mar).
    """
    df = mis_stats[
        (mis_stats["Program"].astype(str).str.strip() == "CWCD") &
        (mis_stats["Description"].isin(INTERVALS))
    ].copy()
    df = df[MONTH_COLS].apply(pd.to_numeric, errors="coerce").fillna(0)  # month-only, numeric
    return df.sum(axis=0)  # total per month across interval rows

def _mis_totals_to_long(mis_totals: pd.Series, fy_start: pd.Timestamp) -> pd.DataFrame:
    """
    Convert Apr..Mar monthly SPI totals to a cumulative-by-month table keyed by YYYY-MM
    for the active FY.
    """
    # Ensure Apr→Mar order and make it cumulative
    ordered = mis_totals.reindex(MONTH_COLS).fillna(0)
    cum_vals = ordered.cumsum()

    fy_year = fy_start.year
    rows = []
    for m in MONTH_COLS:
        year = fy_year if MONTH_NUM[m] >= 4 else fy_year + 1
        ym = f"{year:04d}-{MONTH_NUM[m]:02d}"
        rows.append({"Through Month": ym, "SPI (Cumulative)": float(cum_vals[m])})
    return pd.DataFrame(rows)


def _tidy_pivot_by_month(merged: pd.DataFrame) -> pd.DataFrame:
    """
    1) Rename 'Through Month' → 'Month'
    2) Drop 'Cumulative Start' & 'Cumulative End'
    3) Pivot so months are the columns; keep chronological month ordering.
    """
    df = merged.rename(columns={"Through Month": "Month"}).drop(columns=["Cumulative Start", "Cumulative End"])
    melted = df.melt(id_vars=["Month"], var_name="Metric", value_name="Value")
    pivoted = melted.pivot(index="Metric", columns="Month", values="Value").reset_index()

    # Order columns chronologically by YYYY-MM
    month_cols_sorted = sorted([c for c in pivoted.columns if c != "Metric"], key=lambda s: pd.Period(s, "M"))
    pivoted = pivoted[["Metric"] + month_cols_sorted]
    return pivoted

# ============================== Public API (used by the app) ==============================

def run() -> pd.DataFrame:
    """
    Build the CWCD pivot (metrics x months) and RETURN it as a DataFrame.
    The main app will render this in the Reports tab Output Table.
    """
    # Paths & file picks
    _ = _pick_latest_census_file(DATA_DIR)  # kept for parity with your flow (not needed below)
    visits_csv = DATA_DIR / "rpt_ProgStaffFinance.csv"
    mis_stats_csv = DATA_DIR / "rpt_MIS_Stats.csv"

    # Read
    visits_dat    = _read_csv_utf8sig(visits_csv)
    mis_stats_dat = _read_csv_utf8sig(mis_stats_csv)

    # Fiscal window
    today = pd.Timestamp.today()
    fy_start, eom_last_month, first_of_last_month = _fiscal_year_start_end(today)

    # Visits: FY window + CWCD
    visits_dat["VisitDate"] = pd.to_datetime(visits_dat["Interaction Start Date"], errors="coerce")
    visits_fy   = visits_dat[(visits_dat["VisitDate"] >= fy_start) & (visits_dat["VisitDate"] <= eom_last_month)].copy()
    visits_cwcd = visits_fy[visits_fy["Program"] == "CWCD"].copy()

    # Cumulative month-by-month (Apr → last month)
    cum_by_month = _build_cumulative_table(visits_cwcd, fy_start, eom_last_month)

    # MIS stats totals per month (Apr..Mar) → long YYYY-MM
    mis_totals = _mis_stats_cwcd_monthly_totals(mis_stats_dat)
    spi_long   = _mis_totals_to_long(mis_totals, fy_start)

    # Merge and pivot
    merged  = cum_by_month.merge(spi_long, on="Through Month", how="left")
    pivoted = _tidy_pivot_by_month(merged)

    return pivoted

# Optional: allow command-line preview for debugging
if __name__ == "__main__":
    df = run()
    print(df)
