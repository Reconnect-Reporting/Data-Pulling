# -*- coding: utf-8 -*-
"""
End-to-end pipeline:
- Load Clean Data (Treat) CSVs
- Build metrics (ind_served_out, visits_out, spi_out, spi_group_out, groups_out)
- Fill a copy of the Excel template (named “CMHA Peel - FAME - {LastMonthName} 2025”)
- Email the filled copy to recreport@reconnect.on.ca

Notes:
- The template is read from TEMPLATE_DIR but NEVER modified or written there.
- The filled workbook is saved to %TEMP%\ReconnectReports\... before emailing.
"""

import os
import re
import glob
import shutil
import tempfile
from pathlib import Path
import numpy as np
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# ---------------------- Paths ----------------------
def _onedrive_root() -> Path:
    # Prefer env vars (most reliable on Windows)
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.environ.get(var)
        if p and Path(p).exists():
            return Path(p)
    # Fallback: first "OneDrive*" folder under user profile
    home = Path.home()
    for child in home.iterdir():
        if child.is_dir() and child.name.startswith("OneDrive"):
            return child
    return home / "OneDrive"

OD = _onedrive_root()

# --- Your variables (unchanged names) ---
DATA_DIR = str(OD / "data" / "Clean Data" / "Treat")
CENSUS_CSV = os.path.join(DATA_DIR, "rpt_Census.csv")
VISITS_CSV = os.path.join(DATA_DIR, "rpt_ProgStaffFinance.csv")

TEMPLATE_DIR = str(OD / "data" / "Report Templates")
TEMPLATE_BASENAME = "CMHA Peel - FAME - Template"

# Output (NOT the template folder)
OUTPUT_DIR = os.path.join(os.environ.get("TEMP", tempfile.gettempdir()), "ReconnectReports")

TO_EMAIL = "recreporting@reconnect.on.ca"
EMAIL_BODY = "Hi,\n\nPlease find the attached CMHA Peel – FAME monthly report.\n\nThanks."
SEND_IMMEDIATELY = True  # set False to open draft instead of sending


# ---------------------- Helpers ----------------------
def parse_mdy(series):
    """Parse MDY-like dates; returns pandas datetime (NaT on failure)."""
    return pd.to_datetime(series, errors="coerce")

def parse_mdy_hms(series):
    """Parse MDY HMS-like datetimes; returns pandas datetime (NaT on failure)."""
    return pd.to_datetime(series, errors="coerce")

def between_inclusive(s, start, end):
    return (s >= pd.Timestamp(start)) & (s <= pd.Timestamp(end))

def find_col(df, candidates, required=True):
    """Return the first existing column name from candidates."""
    for c in candidates:
        if c in df.columns:
            return c
    if required:
        raise KeyError(f"None of the expected columns found: {candidates}")
    return None

def read_csv_robust(path):
    """Try UTF-8-SIG, then UTF-8, then latin-1."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            continue
    return pd.read_csv(path)

def to_scalar(val):
    """Sum Series/arrays; coerce to number; fall back to 0."""
    if isinstance(val, (pd.Series, pd.Index, list, tuple)):
        v = pd.to_numeric(pd.Series(val), errors='coerce').fillna(0).sum()
    else:
        try:
            v = float(val)
        except Exception:
            try:
                v = pd.to_numeric(val, errors='coerce')
                v = float(v.sum()) if hasattr(v, "sum") else float(v)
            except Exception:
                v = 0.0
    return int(v) if float(v).is_integer() else v


# ---------------------- Main ----------------------
def main():
    # ----- Read inputs -----
    adt_dat = read_csv_robust(CENSUS_CSV)
    visits_dat = read_csv_robust(VISITS_CSV)

    # ----- Column resolution (Clean Data may differ) -----
    # Census (ADT)
    CN_CLIENT   = find_col(adt_dat, ["ClientName", "Client Name"])
    CN_ID       = find_col(adt_dat, ["ID", "ClientID", "Client ID", "PersonID"])
    CN_AGE      = find_col(adt_dat, ["Age", "ClientAge", "Age (Years)"])
    CN_PROGRAM  = find_col(adt_dat, ["Program", "Program Name"])
    CN_STAFF    = find_col(adt_dat, ["Staff", "Primary Worker", "Worker"])
    CN_CITY     = find_col(adt_dat, ["AddressCity", "City", "Address City", "City/Town"])
    CN_ADMIT    = find_col(adt_dat, ["AdmitDate", "Admission Date", "Admit Date", "DateAdmitted"])
    CN_DISCH    = find_col(adt_dat, ["DischargeDate", "Discharge Date", "Date Discharged", "Exit Date"], required=False)

    # Visits (Clean Data names you showed)
    VS_CLIENT   = find_col(visits_dat, ["ClientName", "Client Name"])
    VS_ID       = find_col(visits_dat, ["ID", "Client ID", "ClientID"])
    VS_VDATE    = find_col(visits_dat, ["Interaction Start Date", "Visit Date", "Date"])
    VS_STIME    = find_col(visits_dat, ["Start Time", "StartTime", "Start"])
    VS_ETIME    = find_col(visits_dat, ["End Time", "EndTime", "End"])
    VS_STAFF    = find_col(visits_dat, ["Staff"])
    VS_TASK     = find_col(visits_dat, ["Task"], required=False)
    VS_PROGRAM  = find_col(visits_dat, ["Program", "Program Name"])
    VS_VTYPE    = find_col(visits_dat, ["Visit Type", "VisitType", "Group Type", "GroupType"])
    VS_NOTE     = find_col(visits_dat, ["Note Date", "NoteDate"], required=False)

    # ----- master list of clients in Peel -----
    adt = (
        adt_dat[[CN_CLIENT, CN_ID, CN_AGE, CN_PROGRAM, CN_STAFF, CN_CITY, CN_ADMIT] + ([CN_DISCH] if CN_DISCH else [])].copy()
        .assign(
            **({CN_DISCH: (lambda df: parse_mdy(df[CN_DISCH]))} if CN_DISCH else {}),
            **{CN_ADMIT:  (lambda df: parse_mdy(df[CN_ADMIT]))},
        )
    )

    # Program + region filters
    cities = [
        "brampton", "mississauga", "caledon", "bolton", "orangeville",
        "burlington", "etobicoke", "milton", "bramp", "brantford"
    ]
    city_pattern = re.compile("|".join(map(re.escape, cities)), flags=re.IGNORECASE)

    today = pd.Timestamp.today().normalize()
    fy_apr1_this_year = pd.Timestamp(year=today.year, month=4, day=1)
    fy_start = fy_apr1_this_year if today >= fy_apr1_this_year else pd.Timestamp(year=today.year - 1, month=4, day=1)

    prog_mask = adt[CN_PROGRAM].isin(["FAMEKids", "FAME"])
    city_mask = adt[CN_CITY].fillna("").astype(str).str.contains(city_pattern)

    if CN_DISCH:
        fy_mask = adt[CN_DISCH].isna() | (adt[CN_ADMIT] >= fy_start) | (adt[CN_DISCH] >= fy_start)
    else:
        fy_mask = (adt[CN_ADMIT] >= fy_start)

    adt = adt[prog_mask & city_mask & fy_mask].copy()

    # Age groups
    def age_group(a):
        try:
            a = float(a)
        except Exception:
            return np.nan
        if a <= 17:
            return "Pediatric"
        if 18 <= a < 65:
            return "Adult"
        if a >= 65:
            return "Elderly"
        return np.nan

    adt["Age Group"] = adt[CN_AGE].apply(age_group)

    # ----- date parameters (last month range) -----
    first_of_this_month = today.replace(day=1)
    som = (first_of_this_month - pd.DateOffset(months=1))          # first day of last month
    eom = first_of_this_month - pd.Timedelta(days=1)               # last day of last month

    # ----- new clients added -----
    ind_served = adt[between_inclusive(adt[CN_ADMIT], som, eom)].copy()
    ind_served["Age Group"] = ind_served["Age Group"].fillna("Unknown")
    ind_served_out = (
        ind_served.groupby("Age Group", dropna=False)
        .size()
        .reset_index(name="n")
    )

    # ----- visits -----
    visits_all = (
        visits_dat.assign(
            ID=pd.to_numeric(visits_dat[VS_ID], errors="coerce"),
            VisitDate=parse_mdy(visits_dat[VS_VDATE]),
        )[["VisitDate", VS_STAFF, VS_VTYPE, VS_PROGRAM, "ID"]]
    )
    visits_all = visits_all[between_inclusive(visits_all["VisitDate"], som, eom)].copy()

    visits_out = (
        adt[[CN_ID, "Age Group"]].rename(columns={CN_ID: "ID"})
        .merge(visits_all, on="ID", how="left")
    )
    visits_out = visits_out[visits_out["VisitDate"].notna()].copy()
    visits_out["Age Group"] = visits_out["Age Group"].fillna("Unknown")
    visits_out = (
        visits_out.groupby(["Age Group", VS_VTYPE], dropna=False)
        .size()
        .reset_index(name="n")
    )
    visits_out = visits_out.rename(columns={VS_VTYPE: "Visit Type"})

    # ----- service provider interactions — individual -----
    spi_all = (
        visits_dat.assign(
            ID=pd.to_numeric(visits_dat[VS_ID], errors="coerce"),
            VisitDate=parse_mdy(visits_dat[VS_VDATE]),
            start_ts=parse_mdy_hms(visits_dat[VS_VDATE].astype(str) + " " + visits_dat[VS_STIME].astype(str)),
            end_ts=parse_mdy_hms(visits_dat[VS_VDATE].astype(str) + " " + visits_dat[VS_ETIME].astype(str)),
        )
        .assign(dTime=lambda df: (df["end_ts"] - df["start_ts"]).dt.total_seconds())
    )
    spi_all = spi_all[
        spi_all["start_ts"].notna() &
        spi_all["end_ts"].notna() &
        spi_all[VS_PROGRAM].isin(["FAMEKids", "FAME"]) &
        between_inclusive(spi_all["VisitDate"], som, eom)
    ].copy()

    spi_out = (
        adt[[CN_ID]].rename(columns={CN_ID: "ID"})
        .merge(spi_all[["ID", "VisitDate", VS_VTYPE, VS_PROGRAM, "dTime"]], on="ID", how="left")
    )
    spi_out = spi_out[spi_out["VisitDate"].notna() & spi_out["dTime"].notna()].copy()

    def interval_individual(seconds):
        if (seconds >= 0) and (seconds <= 1800):    return "5-30 min"
        if (seconds >= 1860) and (seconds <= 3600): return "31-1 hr"
        if (seconds > 3600) and (seconds <= 7200):  return "1-2 hrs"
        if (seconds > 7200) and (seconds <= 18000): return "2-5 hrs"
        if (seconds > 18000):                       return "5 or more hrs"
        return "No Interval"

    spi_out["Interval"] = spi_out["dTime"].apply(interval_individual)
    spi_out = (
        spi_out.groupby("Interval", dropna=False)
        .size()
        .reset_index(name="n")
    )

    # ----- service provider interactions — group (Task-based) -----
    GROUP_TASKS = ["FAME Support Group", "FAMEkids Group"]

    spi_group_all = (
        visits_dat.assign(
            ID=pd.to_numeric(visits_dat[VS_ID], errors="coerce"),
            VisitDate=parse_mdy(visits_dat[VS_VDATE]),
            start_ts=parse_mdy_hms(visits_dat[VS_VDATE].astype(str) + " " + visits_dat[VS_STIME].astype(str)),
            end_ts=parse_mdy_hms(visits_dat[VS_VDATE].astype(str) + " " + visits_dat[VS_ETIME].astype(str)),
        )
        .assign(dTime=lambda df: (df["end_ts"] - df["start_ts"]).dt.total_seconds())
    )[["ID", "VisitDate", VS_TASK, VS_PROGRAM, "dTime"]]

    task_norm = spi_group_all[VS_TASK].astype(str).str.strip() if VS_TASK else pd.Series([], dtype=str)
    spi_group_all = spi_group_all[
        task_norm.isin(GROUP_TASKS) &
        between_inclusive(spi_group_all["VisitDate"], som, eom)
    ].copy()

    spi_group_out = (
        adt[[CN_ID]].rename(columns={CN_ID: "ID"})
        .merge(spi_group_all, on="ID", how="left")
    )
    spi_group_out = spi_group_out[
        spi_group_out["VisitDate"].notna() &
        spi_group_out["dTime"].notna()
    ].copy()

    def interval_group(seconds):
        if (seconds >= 300) and (seconds <= 1800):  return "5-30 min"
        if (seconds >= 1860) and (seconds <= 3600): return "31-1 hr"
        if (seconds > 3600) and (seconds <= 7200):  return "1-2 hrs"
        if (seconds > 7200) and (seconds <= 18000): return "2-5 hrs"
        if (seconds > 18000):                       return "5 or more hrs"
        return np.nan

    spi_group_out["Interval"] = spi_group_out["dTime"].apply(interval_group)
    spi_group_out = (
        spi_group_out.groupby(["Interval", "VisitDate"], dropna=False)
        .size()
        .reset_index(name="n")
    )

    # ----- group sessions (raw rows; Task-based) -----
    groups_all = (
        visits_dat.assign(
            ID=pd.to_numeric(visits_dat[VS_ID], errors="coerce"),
            VisitDate=parse_mdy(visits_dat[VS_VDATE])
        )[["ID", "VisitDate", VS_TASK, VS_PROGRAM]]
    )
    groups_all = groups_all[
        between_inclusive(groups_all["VisitDate"], som, eom) &
        groups_all[VS_TASK].astype(str).str.strip().isin(GROUP_TASKS)
    ].copy()

    groups_out = (
        adt[[CN_ID, "Age Group", CN_PROGRAM]].rename(columns={CN_ID: "ID"})
        .merge(groups_all, on="ID", how="left", suffixes=("", "_visit"))
    )
    groups_out = groups_out[groups_out["VisitDate"].notna()].copy()

    # ----- KPI variables -----
    # Face-to-face (Video/FaceOffice)
    visits_face_to_face_elderly = visits_out.loc[
        (visits_out['Age Group'] == 'Elderly') &
        (visits_out['Visit Type'].isin(['FaceOffice', 'Video']))
    , 'n'].sum()

    visits_face_to_face_adult = visits_out.loc[
        (visits_out['Age Group'] == 'Adult') &
        (visits_out['Visit Type'].isin(['FaceOffice', 'Video']))
    , 'n'].sum()

    visits_face_to_face_Pediatric = visits_out.loc[
        (visits_out['Age Group'] == 'Pediatric') &
        (visits_out['Visit Type'].isin(['FaceOffice', 'Video']))
    , 'n'].sum()

    # Unknowns
    visits_face_to_face_Age_Unknown = visits_out.loc[
        (visits_out['Age Group'].eq('Unknown')) &
        (visits_out['Visit Type'].isin(['FaceOffice', 'Video']))
    , 'n'].sum()

    # Phone/Email
    visits_email_telephone_elderly = visits_out.loc[
        (visits_out['Age Group'] == 'Elderly') &
        (visits_out['Visit Type'].isin(['Phone', 'Email']))
    , 'n'].sum()

    visits_email_telephone_adult = visits_out.loc[
        (visits_out['Age Group'] == 'Adult') &
        (visits_out['Visit Type'].isin(['Phone', 'Email']))
    , 'n'].sum()

    visits_email_telephone_Pediatric = visits_out.loc[
        (visits_out['Age Group'] == 'Pediatric') &
        (visits_out['Visit Type'].isin(['Phone', 'Email']))
    , 'n'].sum()

    visits_email_telephone_Age_unknown = visits_out.loc[
        (visits_out['Age Group'].eq('Unknown')) &
        (visits_out['Visit Type'].isin(['Phone', 'Email']))
    , 'n'].sum()

    # Individuals served
    Indiv_Served_Elderly = ind_served_out.loc[ind_served_out['Age Group'] == 'Elderly', 'n'].sum()
    Indiv_Served_Adult = ind_served_out.loc[ind_served_out['Age Group'] == 'Adult', 'n'].sum()
    Indiv_Served_Pediatric = ind_served_out.loc[ind_served_out['Age Group'] == 'Pediatric', 'n'].sum()
    Indiv_Served_Age_Unknown = ind_served_out.loc[ind_served_out['Age Group'].eq('Unknown'), 'n'].sum()

    # Groups
    Group_Part_Reg_Clients = pd.to_numeric(spi_group_out['n'], errors='coerce').fillna(0).sum()
    Num_of_group_Sessions = len(spi_group_out)

    # SPI (individual) duration bins (Series → summed later)
    spi_5_30_min    = spi_out.loc[spi_out['Interval'] == '5-30 min','n']
    spi_30_1_hour   = spi_out.loc[spi_out['Interval'] == '31-1 hr','n']
    spi_1_2_hour    = spi_out.loc[spi_out['Interval'] == '1-2 hrs','n']
    spi_2_5_hour    = spi_out.loc[spi_out['Interval'] == '2-5 hrs','n']
    spi_5_more_hour = spi_out.loc[spi_out['Interval'] == '5 or more hrs','n']

    # SPI (group) duration bins (exact counts as ints)
    _bins = ["5-30 min", "31-1 hr", "1-2 hrs", "2-5 hrs", "5 or more hrs"]
    interval_clean = spi_group_out["Interval"].astype("string").str.strip()
    _counts = interval_clean.value_counts().reindex(_bins, fill_value=0)
    spgi_5_30_min      = int(_counts["5-30 min"])
    spgi_31_1_hr       = int(_counts["31-1 hr"])
    spgi_1_2_hrs       = int(_counts["1-2 hrs"])
    spgi_2_5_hrs       = int(_counts["2-5 hrs"])
    spgi_5_or_more_hrs = int(_counts["5 or more hrs"])

    # ---------------------- Fill a copy of the Excel template (NOT in template folder) ----------------------
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    candidates = (
        glob.glob(os.path.join(TEMPLATE_DIR, TEMPLATE_BASENAME + "*.xlsm")) +
        glob.glob(os.path.join(TEMPLATE_DIR, TEMPLATE_BASENAME + "*.xlsx"))
    )
    if not candidates:
        raise FileNotFoundError(f"No Excel template named like '{TEMPLATE_BASENAME}*.xlsm/.xlsx' in {TEMPLATE_DIR}")
    template_path = candidates[0]
    keep_vba = template_path.lower().endswith(".xlsm")

    report_month_name = (today.replace(day=1) - pd.DateOffset(months=1)).strftime("%B")  # e.g., "July"
    base_template_ext = os.path.splitext(template_path)[1]  # keep .xlsx/.xlsm
    out_name = f"CMHA Peel - FAME - {report_month_name} 2025{base_template_ext}"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    # Load template (read-only source), then save the FILLED copy to OUTPUT_DIR
    wb = load_workbook(template_path, keep_vba=keep_vba, data_only=False)
    ws = wb.active  # change to wb['YourSheetName'] if your template uses a specific sheet

    cell_to_var = {
        # Face-to-face (Video/FaceOffice)
        "C6":  "visits_face_to_face_elderly",
        "C7":  "visits_face_to_face_adult",
        "C8":  "visits_face_to_face_Pediatric",
        "C9":  "visits_face_to_face_Age_Unknown",

        # Phone/Email
        "C10": "visits_email_telephone_elderly",
        "C11": "visits_email_telephone_adult",
        "C12": "visits_email_telephone_Pediatric",
        "C13": "visits_email_telephone_Age_unknown",

        # Individuals served
        "C15": "Indiv_Served_Elderly",
        "C16": "Indiv_Served_Adult",
        "C17": "Indiv_Served_Pediatric",
        "C18": "Indiv_Served_Age_Unknown",

        # Groups
        "C26": "Group_Part_Reg_Clients",
        "C27": "Num_of_group_Sessions",

        # SPI (individual) duration bins C36–C40
        "C36": "spi_5_30_min",
        "C37": "spi_30_1_hour",
        "C38": "spi_1_2_hour",
        "C39": "spi_2_5_hour",
        "C40": "spi_5_more_hour",

        # SPI (group) duration bins C53–C57
        "C53": "spgi_5_30_min",
        "C54": "spgi_31_1_hr",
        "C55": "spgi_1_2_hrs",
        "C56": "spgi_2_5_hrs",
        "C57": "spgi_5_or_more_hrs",
    }

    locs = locals()
    for cell, varname in cell_to_var.items():
        ws[cell] = to_scalar(locs.get(varname, 0))

    wb.save(out_path)
    wb.close()

    # ---------------------- Email the file ----------------------
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = TO_EMAIL
    mail.Subject = f"CMHA Peel – FAME Monthly Report – {report_month_name} 2025"
    mail.Body = EMAIL_BODY
    mail.Attachments.Add(out_path)
    if SEND_IMMEDIATELY:
        mail.Send()
    else:
        mail.Display()


if __name__ == "__main__":
    main()
