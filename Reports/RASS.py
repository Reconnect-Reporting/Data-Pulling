# Reports/UHN_GeriPsych.py
# -*- coding: utf-8 -*-
# Preview-first: run() returns a 2-col DataFrame for the app's Output Table
# run_dict() returns the original dict of metrics if you need it elsewhere.

import getpass
from pathlib import Path
from typing import Optional
import pandas as pd

# ---------------------- Config ----------------------
USER_NAME = getpass.getuser()
DATA_DIR = Path(
    fr"C:\Users\{USER_NAME}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat"
)
PROGRAM_FILTER = "RASS"   # substring match (case-insensitive)

# ---------------------- Helpers ----------------------
def _read_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")

def _parse_date_multi(series: pd.Series, formats=("%m/%d/%Y", "%Y-%m-%d")) -> pd.Series:
    """Try multiple date formats; fallback to pandas' parser."""
    s = pd.to_datetime(series, format=formats[0], errors="coerce")
    if s.isna().any():
        mask = s.isna()
        for fmt in formats[1:]:
            if not mask.any():
                break
            s2 = pd.to_datetime(series[mask], format=fmt, errors="coerce")
            s.loc[mask] = s2
            mask = s.isna()
    fallback = pd.to_datetime(series, errors="coerce")
    s = s.fillna(fallback)
    return s

def _prog_contains(series: pd.Series, needle: str) -> pd.Series:
    return series.astype(str).str.contains(needle, case=False, na=False)

def _ensure_id_from_clientname(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure an integer ID column exists; extract from ClientName if missing."""
    if "ID" not in df.columns or (df["ID"].isna().all() if "ID" in df.columns else True):
        if "ClientName" in df.columns:
            df["ID"] = pd.to_numeric(df["ClientName"].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    return df

def _first_event_mean_scalar(
    df: pd.DataFrame,
    event_mask: pd.Series,
    date_col: str,            # event date column to minimize (e.g., "Note Date" or "Interaction Start Date")
    days_col: str,            # precomputed days difference col to average
) -> Optional[float]:
    """Mean (days) from admission to FIRST event across clients. Returns rounded float or None."""
    if df.empty or date_col not in df.columns or days_col not in df.columns:
        return None
    if not isinstance(event_mask, pd.Series):
        return None
    if len(event_mask) != len(df):
        event_mask = pd.Series([False]*len(df), index=df.index)

    sub = df.loc[event_mask & df[date_col].notna(), ["ID", date_col, days_col]].dropna(subset=["ID", days_col])
    if sub.empty:
        return None
    try:
        idx = sub.groupby("ID", dropna=False)[date_col].idxmin()
    except Exception:
        return None
    first_rows = sub.loc[idx]
    if first_rows.empty:
        return None
    val = first_rows[days_col].mean()
    return None if pd.isna(val) else round(float(val), 1)

# ---------------------- Core compute (returns dict) ----------------------
def run_dict() -> dict:
    """
    Returns:
      - Numbers (scalars):
          * Number of Referrals
          * Number of Active Clients
          * Number of Discharged Clients
          * Avg days: admission → first note
          * Avg days: admission → first NF2F
          * Avg days: admission → first F2F
          * Avg days: admission → discharge
          * F2F No Shows
          * NF2F No Shows
      - Table:
          * final_out_visits_Total (No-shows only, renamed)
    """
    # ---------- Read ----------
    referrals = _read_csv(DATA_DIR / "rpt_Referrals_DT.csv")
    census    = _read_csv(DATA_DIR / "rpt_Census.csv")
    notes     = _read_csv(DATA_DIR / "rpt_ProgStaffFinance.csv")

    # ---------- Parse & filter ----------
    # Census
    if "Program" in census.columns:
        census = census[_prog_contains(census["Program"], PROGRAM_FILTER)].copy()
    if "AdmitDate" in census.columns:
        census["AdmitDate"] = _parse_date_multi(census["AdmitDate"])
    if "DischargeDate" in census.columns:
        census["DischargeDate"] = _parse_date_multi(census["DischargeDate"])
    census = _ensure_id_from_clientname(census)

    # Referrals
    for col in ("ReferralDate", "DateDecisionMade"):
        if col in referrals.columns:
            referrals[col] = _parse_date_multi(referrals[col])
    if "Program" in referrals.columns:
        referrals = referrals[_prog_contains(referrals["Program"], PROGRAM_FILTER)].copy()

    # Notes — keep your existing column names (including "Note Date")
    notes = _ensure_id_from_clientname(notes)
    if "Program" in notes.columns:
        notes = notes[_prog_contains(notes["Program"], PROGRAM_FILTER)].copy()

    # Parse dates if present (new column names)
    if "Interaction Start Date" in notes.columns:
        notes["Interaction Start Date"] = _parse_date_multi(notes["Interaction Start Date"])
    if "Note Date" in notes.columns:
        notes["Note Date"] = _parse_date_multi(notes["Note Date"])

    # NF2F remap via Task
    if "Task" in notes.columns:
        NF2F = {"Phone call (Direct Contact)", "Email/Mail/Text/Fax"}
        notes.loc[notes["Task"].isin(NF2F), "Task"] = "NF2F"

    # ---------- Counts ----------
    num_of_referrals = int(len(referrals))
    if "DischargeDate" in census.columns:
        num_of_active_client     = int(census["DischargeDate"].isna().sum())
        num_of_discharged_client = int(census["DischargeDate"].notna().sum())
    else:
        num_of_active_client = 0
        num_of_discharged_client = 0

    # ---------- Merge for durations ----------
    left_cols  = [c for c in ["ClientName", "ID", "AdmitDate", "DischargeDate"] if c in census.columns]
    right_cols = [c for c in ["ID", "Program", "Note Date", "Interaction Start Date", "Task"] if c in notes.columns]

    if "ID" in census.columns and "ID" in notes.columns:
        merged = census.loc[:, left_cols].merge(notes.loc[:, right_cols], on="ID", how="inner")
    else:
        merged = pd.DataFrame(columns=list(set(left_cols + right_cols)))

    if not merged.empty:
        if {"Note Date", "AdmitDate"}.issubset(merged.columns):
            merged["AdmissionDate_NoteDate"] = (merged["Note Date"] - merged["AdmitDate"]).dt.days
            # keep parity with earlier logic: interaction delta from Note Date
            merged["AdmissionDate_InteractionDate"] = (merged["Note Date"] - merged["AdmitDate"]).dt.days
        else:
            merged["AdmissionDate_NoteDate"] = pd.NA
            merged["AdmissionDate_InteractionDate"] = pd.NA

        if {"DischargeDate", "AdmitDate"}.issubset(merged.columns):
            merged["AdmissionDate_DischargeDate"] = (merged["DischargeDate"] - merged["AdmitDate"]).dt.days
        else:
            merged["AdmissionDate_DischargeDate"] = pd.NA
    else:
        for col in ["AdmissionDate_NoteDate","AdmissionDate_InteractionDate","AdmissionDate_DischargeDate"]:
            if col not in merged.columns:
                merged[col] = pd.Series(dtype="float64")

    # ---------- Averages (numbers only) ----------
    # Build safe boolean masks (Series) only when columns exist
    mask_note_exists = ("Note Date" in merged.columns)
    mask_task_exists = ("Task" in merged.columns)
    mask_inter_exists = ("Interaction Start Date" in merged.columns)

    event_mask_first_note = merged["Note Date"].notna() if mask_note_exists else pd.Series([False]*len(merged), index=merged.index)
    event_mask_nf2f = (merged["Task"] == "NF2F") if mask_task_exists else pd.Series([False]*len(merged), index=merged.index)
    event_mask_f2f  = (merged["Task"] == "Home/Community Visit") if mask_task_exists else pd.Series([False]*len(merged), index=merged.index)

    avg_days_first_note = _first_event_mean_scalar(
        merged, event_mask=event_mask_first_note, date_col="Note Date", days_col="AdmissionDate_NoteDate",
    )
    avg_days_first_nf2f = _first_event_mean_scalar(
        merged, event_mask=event_mask_nf2f if mask_inter_exists else pd.Series([False]*len(merged), index=merged.index),
        date_col="Interaction Start Date", days_col="AdmissionDate_InteractionDate",
    )
    avg_days_first_f2f = _first_event_mean_scalar(
        merged, event_mask=event_mask_f2f if mask_inter_exists else pd.Series([False]*len(merged), index=merged.index),
        date_col="Interaction Start Date", days_col="AdmissionDate_InteractionDate",
    )

    # Admission → Discharge (earliest discharge per ID)
    avg_days_admin_dis = None
    if not merged.empty and {"DischargeDate", "AdmissionDate_DischargeDate"}.issubset(merged.columns):
        sub = merged.loc[:, ["ID", "DischargeDate", "AdmissionDate_DischargeDate"]].dropna()
        if not sub.empty:
            try:
                idx = sub.groupby("ID", dropna=False)["DischargeDate"].idxmin()
                first_dis = sub.loc[idx]
                val = first_dis["AdmissionDate_DischargeDate"].mean()
                avg_days_admin_dis = None if pd.isna(val) else round(float(val), 1)
            except Exception:
                avg_days_admin_dis = None

    # ---------- No-shows (scalars) ----------
    # F2F: "No Show";  NF2F: "Phone Call (No Contact)"
    f2f_no_shows = 0
    nf2f_no_shows = 0
    if not notes.empty and "Task" in notes.columns:
        f2f_no_shows  = int((notes["Task"] == "No Show").sum())
        nf2f_no_shows = int((notes["Task"] == "Phone Call (No Contact)").sum())

    # ---------- final_out_visits_Total (No-shows only, renamed) ----------
    if not notes.empty and {"Program", "Task"}.issubset(notes.columns):
        mapping = {
            "No Show": "F2F No shows",
            "Phone Call (No Contact)": "NF2F No shows",
        }
        ns = notes[notes["Task"].isin(mapping.keys())].copy()
        ns["Task"] = ns["Task"].map(mapping)
        final_out_visits_Total = (
            ns.groupby(["Program", "Task"], dropna=False)
              .size()
              .reset_index(name="count")
              .sort_values(["Program", "Task"], kind="stable")
        )
    else:
        final_out_visits_Total = pd.DataFrame(columns=["Program", "Task", "count"])

    # ---------- Return ----------
    return {
        "Number of Referrals": num_of_referrals,
        "Number of Active Clients": num_of_active_client,
        "Number of Discharged Clients": num_of_discharged_client,
        "Avg days: admission → first note": avg_days_first_note,
        "Avg days: admission → first NF2F": avg_days_first_nf2f,
        "Avg days: admission → first F2F": avg_days_first_f2f,
        "Avg days: admission → discharge": avg_days_admin_dis,
        "F2F No Shows": f2f_no_shows,
        "NF2F No Shows": nf2f_no_shows,
        "final_out_visits_Total": final_out_visits_Total,  # optional for UI (not shown in preview table)
    }

# ---------------------- Preview wrapper (returns DataFrame) ----------------------
def _to_preview_df(summary: dict) -> pd.DataFrame:
    """Convert the dict of metrics into a 2-col DataFrame (Metric | Value)."""
    rows = []
    for k, v in summary.items():
        if k == "final_out_visits_Total":   # skip the detail table in the preview
            continue
        rows.append({"Metric": k, "Value": v})
    return pd.DataFrame(rows, columns=["Metric", "Value"])

def run() -> pd.DataFrame:
    """Entry used by the app: returns a DataFrame so it renders in Output Table."""
    return _to_preview_df(run_dict())

# ---------------------- CLI preview ----------------------
if __name__ == "__main__":
    # When run directly: print the numbers (dict) and also show the preview table
    outs = run_dict()
    def show(x): return "N/A" if x is None else x

    print(f"Number of Referrals: {show(outs['Number of Referrals'])}")
    print(f"Number of Active Clients: {show(outs['Number of Active Clients'])}")
    print(f"Number of Discharged Clients: {show(outs['Number of Discharged Clients'])}")
    print(f"F2F No Shows: {outs['F2F No Shows']}")
    print(f"NF2F No Shows: {outs['NF2F No Shows']}")
    print(f"Avg days: admission → first note: {show(outs['Avg days: admission → first note'])}")
    print(f"Avg days: admission → first NF2F: {show(outs['Avg days: admission → first NF2F'])}")
    print(f"Avg days: admission → first F2F: {show(outs['Avg days: admission → first F2F'])}")
    print(f"Avg days: admission → discharge: {show(outs['Avg days: admission → discharge'])}")

    print("\n[Preview DataFrame]")
    print(_to_preview_df(outs))
