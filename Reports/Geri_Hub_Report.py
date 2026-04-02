# Reports/GeriHub_Summary.py

import os
from pathlib import Path
from typing import Optional
import pandas as pd

# ---------- Config ----------
USER = os.getlogin()
Caredove_DIR = Path(fr"C:\Users\{USER}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Caredove")
Alaya_DIR    = Path(fr"C:\Users\{USER}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\AlayaCare")

CD_BASENAME    = "CaredoveBookingsReport_Reconnect"
TARGET_SERVICE = "Geri Hub Referrals for Social Work Case Management Program"

# ---------- Helpers ----------
def _find_excel(folder: Path, base: str) -> Path:
    for ext in (".xlsx", ".xls"):
        p = folder / f"{base}{ext}"
        if p.exists():
            return p
    for p in sorted(list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))):
        if p.stem.startswith(base):
            return p
    raise FileNotFoundError(f"Excel starting with '{base}' not found in {folder}")

def _read_excel_any(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    required = {"Service Selected", "Date/Time Submitted", "Reference #"}
    for sh in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sh, dtype=str)
        if required.issubset(df.columns):
            return df
    return pd.read_excel(path, sheet_name=0, dtype=str)

def newest_any(prefix: str, base: Path, exts=(".csv", ".xlsx", ".xls")) -> Path:
    files = []
    for ext in exts:
        files.extend(base.glob(f"{prefix}*{ext}"))
    files = sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No '{prefix}*{exts}' found in {base}")
    return files[0]

def read_any(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc, keep_default_na=False)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    return pd.read_excel(path, dtype=str)

def prev_full_month_window(today: Optional[pd.Timestamp] = None):
    if today is None:
        today = pd.Timestamp.today()
    month_start = pd.Timestamp(today.year, today.month, 1)
    prev_end = month_start - pd.Timedelta(days=1)
    prev_start = pd.Timestamp(prev_end.year, prev_end.month, 1)
    return prev_start, prev_end, month_start

# ---------- Core builder ----------
def _build_tables():
    # --- Caredove: filter to service + previous full month ---
    cd_path = _find_excel(Caredove_DIR, CD_BASENAME)
    cd_df   = _read_excel_any(cd_path)

    svc_col        = "Service Selected"
    date_col       = "Date/Time Submitted"
    ref_col        = "Reference #"
    requester_col  = "Requester Name"

    cd_df = cd_df[cd_df[svc_col].astype(str).str.strip() == TARGET_SERVICE].copy()

    prev_start, prev_end, month_start = prev_full_month_window()
    cd_df["__submitted"] = pd.to_datetime(cd_df[date_col], errors="coerce", utc=True).dt.tz_convert(None)
    mask_cd = (cd_df["__submitted"] >= prev_start) & (cd_df["__submitted"] < month_start)
    cd_prev = cd_df.loc[mask_cd].copy().fillna("N/A")

    # Number of Reference
    count_unique_refs = cd_prev[ref_col].astype(str).str.strip().nunique()
    table_refs = pd.DataFrame(
        [{"Name": "", "Count": count_unique_refs, "Section": "Number of Reference"}],
        columns=["Section", "Name", "Count"]
    )[["Section", "Name", "Count"]]

    # Requester Name counts
    if requester_col in cd_prev.columns:
        table_requester = (
            cd_prev.groupby([requester_col]).size().reset_index(name="Count")
            .rename(columns={requester_col: "Name"})
            .assign(Section="Requester Name")
        )[["Section", "Name", "Count"]]
    else:
        table_requester = pd.DataFrame(columns=["Section", "Name", "Count"])

    # --- AlayaCare Visits: Geriatric Hub + previous full month ---
    visits_path = newest_any("Visits", Alaya_DIR)
    visits_df   = read_any(visits_path)

    g_hub = visits_df[
    visits_df["Service Description"]
        .astype(str)
        .str.contains("geri", case=False, na=False)].copy()
    
    g_hub["__Date"] = pd.to_datetime(g_hub["Date"], errors="coerce", utc=True).dt.tz_convert(None)
    mask_ac = (g_hub["__Date"] >= prev_start) & (g_hub["__Date"] < month_start)
    g_hub_prev = g_hub.loc[mask_ac].copy()

    # Number of Visit (unique visits)
    g_hub_visit = int(g_hub_prev["Visit Visit ID"].nunique())
    table_num_visit = pd.DataFrame(
        [{"Name": "", "Count": g_hub_visit, "Section": "Number of Visit"}],
        columns=["Section", "Name", "Count"]
    )[["Section", "Name", "Count"]]

    # Individuals Served (unique clients)
    g_hub_individual_served = int(g_hub_prev["AlaycareID"].nunique())
    table_individuals = pd.DataFrame(
        [{"Name": "", "Count": g_hub_individual_served, "Section": "Individuals Served"}],
        columns=["Section", "Name", "Count"]
    )[["Section", "Name", "Count"]]

    # Combine for Output tab
    combined_out = pd.concat(
        [table_refs, table_requester, table_num_visit, table_individuals],
        ignore_index=True
    )[["Section", "Name", "Count"]]

    return table_refs, table_requester, table_num_visit, table_individuals, combined_out

# ---------- Public API ----------
def run() -> pd.DataFrame:
    """
    Returns ONE combined DataFrame:
      Section | Name | Count
    """
    *_tables, combined_out = _build_tables()
    return combined_out

def main():
    table_refs, table_requester, table_num_visit, table_individuals, combined_out = _build_tables()

    with pd.option_context("display.max_columns", None, "display.width", 160):
        print("\n== Number of Reference ==")
        print(table_refs)

        print("\n== Requester Name ==")
        print(table_requester)

        print("\n== Number of Visit ==")
        print(table_num_visit)

        print("\n== Individuals Served ==")
        print(table_individuals)

        print("\n== Combined (for Output tab) ==")
        print(combined_out)

if __name__ == "__main__":
    main()
