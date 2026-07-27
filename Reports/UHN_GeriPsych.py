# Reports/UHN_GeriPsych.py
# -*- coding: utf-8 -*-

import getpass
from pathlib import Path
from typing import Optional
import pandas as pd

# ---------------------- Config ----------------------
USER_NAME = getpass.getuser()
DATA_DIR = Path(
    fr"C:\Users\{USER_NAME}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat"
)
PROGRAM_FILTER = "UHN Geriatric Psychiatry"

# ONLY for avg-day indicators
ADMISSION_CUTOFF = pd.Timestamp("2026-01-01")

# ---------------------- Helpers ----------------------
def _read_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def _parse_date_multi(series: pd.Series, formats=("%m/%d/%Y", "%Y-%m-%d")) -> pd.Series:
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
    return s.fillna(fallback)


def _prog_contains(series: pd.Series, needle: str) -> pd.Series:
    return series.astype(str).str.contains(needle, case=False, na=False)


def _ensure_id_from_clientname(df: pd.DataFrame) -> pd.DataFrame:
    if "ID" not in df.columns or df["ID"].isna().all():
        if "ClientName" in df.columns:
            df["ID"] = pd.to_numeric(
                df["ClientName"].astype(str).str.extract(r"(\d+)")[0],
                errors="coerce",
            )
    return df


def _first_event_mean_scalar(
    df: pd.DataFrame,
    event_mask: pd.Series,
    date_col: str,
    days_col: str,
) -> Optional[float]:

    if df.empty or date_col not in df.columns or days_col not in df.columns:
        return None

    if len(event_mask) != len(df):
        event_mask = pd.Series(False, index=df.index)

    sub = df.loc[
        event_mask
        & df[date_col].notna()
        & df[days_col].notna(),
        ["ID", date_col, days_col]
    ].dropna()

    if sub.empty:
        return None

    try:
        idx = sub.groupby("ID")[date_col].idxmin()
    except Exception:
        return None

    first_rows = sub.loc[idx]
    val = first_rows[days_col].mean()

    return None if pd.isna(val) else round(float(val), 1)


# ---------------------- Core compute ----------------------
def run_dict() -> dict:

    # ---------- Read ----------
    referrals = _read_csv(DATA_DIR / "rpt_Referrals_DT.csv")
    census = _read_csv(DATA_DIR / "rpt_Census.csv")
    notes = _read_csv(DATA_DIR / "rpt_ProgStaffFinance.csv")

    # ---------- Census ----------
    if "Program" in census.columns:
        census = census[_prog_contains(census["Program"], PROGRAM_FILTER)].copy()

    if "AdmitDate" in census.columns:
        census["AdmitDate"] = _parse_date_multi(census["AdmitDate"])

    if "DischargeDate" in census.columns:
        census["DischargeDate"] = _parse_date_multi(census["DischargeDate"])

    census = _ensure_id_from_clientname(census)

    # ---------- Referrals ----------
    for col in ("ReferralDate", "DateDecisionMade"):
        if col in referrals.columns:
            referrals[col] = _parse_date_multi(referrals[col])

    if "Program" in referrals.columns:
        referrals = referrals[_prog_contains(referrals["Program"], PROGRAM_FILTER)].copy()

    # ---------- Notes ----------
    notes = _ensure_id_from_clientname(notes)

    if "Program" in notes.columns:
        notes = notes[_prog_contains(notes["Program"], PROGRAM_FILTER)].copy()

    if "Interaction Start Date" in notes.columns:
        notes["Interaction Start Date"] = _parse_date_multi(notes["Interaction Start Date"])

    if "Note Date" in notes.columns:
        notes["Note Date"] = _parse_date_multi(notes["Note Date"])

    if "Task" in notes.columns:
        NF2F = {"Phone call (Direct Contact)", "Email/Mail/Text/Fax"}
        notes.loc[notes["Task"].isin(NF2F), "Task"] = "NF2F"

    # ---------- Counts (UNFILTERED) ----------
    num_of_referrals = int(len(referrals))

    if "DischargeDate" in census.columns:
        num_of_active_client = int(census["DischargeDate"].isna().sum())
        num_of_discharged_client = int(census["DischargeDate"].notna().sum())
    else:
        num_of_active_client = 0
        num_of_discharged_client = 0

    # ---------- Merge ----------
    left_cols = [
        c for c in ["ClientName", "ID", "AdmitDate", "DischargeDate"]
        if c in census.columns
    ]

    right_cols = [
        c for c in ["ID", "Program", "Note Date", "Interaction Start Date", "Task"]
        if c in notes.columns
    ]

    merged_all = census[left_cols].merge(
        notes[right_cols],
        on="ID",
        how="inner"
    )

    if not merged_all.empty:

        # Make sure contact date is based on Interaction Start Date
        # This is the actual contact date used to find first NF2F / F2F.
        if {"Interaction Start Date", "AdmitDate"}.issubset(merged_all.columns):
            merged_all["AdmissionDate_InteractionDate"] = (
                merged_all["Interaction Start Date"] - merged_all["AdmitDate"]
            ).dt.days

            # Important:
            # Only keep contact rows where Interaction Start Date is later than AdmitDate.
            # This prevents contacts before admission or same-day invalid rows from being used
            # for first NF2F / first F2F average calculations.
            merged_all["Contact_After_Admission"] = (
                merged_all["Interaction Start Date"].notna()
                & merged_all["AdmitDate"].notna()
                & (merged_all["Interaction Start Date"] > merged_all["AdmitDate"])
            )
        else:
            merged_all["AdmissionDate_InteractionDate"] = pd.NA
            merged_all["Contact_After_Admission"] = False

        if {"DischargeDate", "AdmitDate"}.issubset(merged_all.columns):
            merged_all["AdmissionDate_DischargeDate"] = (
                merged_all["DischargeDate"] - merged_all["AdmitDate"]
            ).dt.days
        else:
            merged_all["AdmissionDate_DischargeDate"] = pd.NA

    else:
        merged_all["AdmissionDate_InteractionDate"] = pd.NA
        merged_all["AdmissionDate_DischargeDate"] = pd.NA
        merged_all["Contact_After_Admission"] = False

    # ---------- Cutoff ONLY for avg indicators ----------
    merged_ind = merged_all[
        merged_all["AdmitDate"].notna()
        & (merged_all["AdmitDate"] >= ADMISSION_CUTOFF)
    ].copy()

    # ---------- Avg days ----------
    # Only include contacts after admission date
    event_mask_nf2f = (
        (merged_ind["Task"] == "NF2F")
        & (merged_ind["Contact_After_Admission"] == True)
    )

    event_mask_f2f = (
        (merged_ind["Task"] == "Home/Community Visit")
        & (merged_ind["Contact_After_Admission"] == True)
    )

    avg_days_first_nf2f = _first_event_mean_scalar(
        merged_ind,
        event_mask_nf2f,
        "Interaction Start Date",
        "AdmissionDate_InteractionDate",
    )

    avg_days_first_f2f = _first_event_mean_scalar(
        merged_ind,
        event_mask_f2f,
        "Interaction Start Date",
        "AdmissionDate_InteractionDate",
    )

    avg_days_admin_dis = None

    if not merged_ind.empty:
        sub = merged_ind[
            ["ID", "DischargeDate", "AdmissionDate_DischargeDate"]
        ].dropna()

        if not sub.empty:
            idx = sub.groupby("ID")["DischargeDate"].idxmin()
            val = sub.loc[idx]["AdmissionDate_DischargeDate"].mean()
            avg_days_admin_dis = None if pd.isna(val) else round(float(val), 1)

    # ---------- No-shows ----------
    f2f_no_shows = int((notes["Task"] == "No Show").sum())
    nf2f_no_shows = int((notes["Task"] == "Phone Call (No Contact)").sum())

    # ---------- Return ----------
    return {
        "Number of Referrals": num_of_referrals,
        "Number of Active Clients": num_of_active_client,
        "Number of Discharged Clients": num_of_discharged_client,
        "Avg days: admission → first NF2F": avg_days_first_nf2f,
        "Avg days: admission → first F2F": avg_days_first_f2f,
        "Avg days: admission → discharge": avg_days_admin_dis,
        "F2F No Shows": f2f_no_shows,
        "NF2F No Shows": nf2f_no_shows,
    }


# ---------------------- Preview wrapper ----------------------
def _to_preview_df(summary: dict) -> pd.DataFrame:

    rows = []
    for k, v in summary.items():
        rows.append({"Metric": k, "Value": v})

    return pd.DataFrame(rows, columns=["Metric", "Value"])


def run() -> pd.DataFrame:
    return _to_preview_df(run_dict())


# ---------------------- CLI ----------------------
if __name__ == "__main__":

    outs = run_dict()

    def show(x):
        return "N/A" if x is None else x

    print(f"Number of Referrals: {show(outs['Number of Referrals'])}")
    print(f"Number of Active Clients: {show(outs['Number of Active Clients'])}")
    print(f"Number of Discharged Clients: {show(outs['Number of Discharged Clients'])}")
    print(f"F2F No Shows: {outs['F2F No Shows']}")
    print(f"NF2F No Shows: {outs['NF2F No Shows']}")
    print(f"Avg days: admission → first NF2F: {show(outs['Avg days: admission → first NF2F'])}")
    print(f"Avg days: admission → first F2F: {show(outs['Avg days: admission → first F2F'])}")
    print(f"Avg days: admission → discharge: {show(outs['Avg days: admission → discharge'])}")

    print("\n[Preview DataFrame]")
    print(_to_preview_df(outs))