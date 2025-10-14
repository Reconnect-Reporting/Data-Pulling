# trapeze_pipeline.py
from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


# ========= utils =========

def get_username() -> str:
    return (
        os.environ.get("USERName")
        or os.environ.get("USERNAME")
        or os.environ.get("UserName")
        or os.getlogin()
    )

def expand_user_placeholder(path_tpl: str, user: str) -> Path:
    """Replace '=' in a Windows path template with the username."""
    return Path(path_tpl.replace("=", user))

def excel_letters(n: int) -> list[str]:
    """A, B, ..., Z, AA, AB, ... up to n columns."""
    out = []
    i = 0
    while len(out) < n:
        s = ""
        x = i
        while True:
            s = chr(ord("A") + (x % 26)) + s
            x = x // 26 - 1
            if x < 0:
                break
        out.append(s)
        i += 1
    return out[:n]

def find_latest_trapeze_file(folder: Path) -> tuple[Path, str | None]:
    """
    Find the most recent Trapeze file by END date in name:
      Trapeze_MM-DD-YYYY_MM-DD-YYYY.(csv|xlsx|xls)
    Returns (path, sheet_name_if_excel_or_None).
    """
    pat = re.compile(
        r"^Trapeze_(\d{2}-\d{2}-\d{4})_(\d{2}-\d{2}-\d{4})\.(csv|xlsx|xls)$",
        re.IGNORECASE,
    )
    best = None
    best_key = None

    for p in folder.glob("Trapeze_*_*.*"):
        m = pat.match(p.name)
        if not m:
            continue
        try:
            end_dt = datetime.strptime(m.group(2), "%m-%d-%Y")
        except ValueError:
            continue
        key = (end_dt, p.stat().st_mtime)
        if best_key is None or key > best_key:
            best, best_key = p, key

    if best is None:
        raise FileNotFoundError(f"No Trapeze_*_*.csv/.xlsx found in {folder}")

    sheet = None
    if best.suffix.lower() in (".xlsx", ".xls"):
        with pd.ExcelFile(best) as xf:
            sheet = xf.sheet_names[0] if xf.sheet_names else None
    return best, sheet


# ========= main pipeline logic =========

def run() -> Path:
    usern = get_username()

    # --- Locate Trapeze folder and load MOST RECENT file ---
    TRAPEZE_DIR_TPL = r"C:\Users\=\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Trapeze"
    TRAPEZE_DIR = expand_user_placeholder(TRAPEZE_DIR_TPL, usern)

    latest_file, excel_sheet = find_latest_trapeze_file(TRAPEZE_DIR)

    # Read as CSV or Excel (first sheet), assign LETTERS as columns
    if latest_file.suffix.lower() == ".csv":
        te = pd.read_csv(latest_file, header=None, dtype="object", encoding="utf-8-sig")
    else:
        te = pd.read_excel(latest_file, header=None, dtype="object", sheet_name=excel_sheet)

    te.columns = excel_letters(te.shape[1])

    # ---- Transformations (mirrors original logic) ----
    I_num = pd.to_numeric(te.get("I"), errors="coerce")
    te = te.loc[I_num == 9].copy()
    te["I"] = "CL"

    te["Stats.Code"] = "a"
    te = te[list(te.columns[:10]) + [te.columns[25]] + list(te.columns[10:25])].copy()

    te["K"] = pd.to_numeric(te["K"], errors="coerce").round(2)
    te.loc[te["X"] == "TAXI- RCH", "G"] = "Performed"

    te.loc[te["X"] == "TAXI- RCH", "X"] = "RCH 2"
    te.loc[te["X"] == "TAXI- WN",  "X"] = "RCH 2"
    te.loc[te["X"] == "TAXI- SF",  "X"] = "RCH 2"
    te.loc[te["X"] == "TAXI - SP", "X"] = "RCH 2"

    te["Q"] = pd.to_numeric(te["Q"], errors="coerce")

    # ADP rules
    mask_adp_base = te["G"].isin(["Performed", "Scheduled", "Arrived"])
    te.loc[(te["M"] == "RCH ADS") & mask_adp_base, "Stats.Code"] = "TRADS!"
    te.loc[(te["M"] == "HU ADS")  & mask_adp_base, "Stats.Code"] = "TRADS!"
    te.loc[(te["M"] == "WT ADS")  & mask_adp_base, "Stats.Code"] = "TRADS!"
    te.loc[(te["M"] == "RCH PRG") & mask_adp_base, "Stats.Code"] = "TRADS!"
    te.loc[(te["M"] == "RCH ADS") & te["G"].isin(["Performed","Unknown"]), "Stats.Code"] = "TRADS!"
    te.loc[(te["M"] == "RCH ADS") & te["G"].isin(["Performed","Scheduled"]) & (te["K"] == 0), "Stats.Code"] = "TRADSFW!"
    te.loc[
        (te["M"] == "RCH ADS")
        & te["G"].isin(["Unknown Cancellation","Cancel At Door","Cancelled Due To Destination Location Being Closed"])
        & te["K"].isin([4, 4.25, 4.35, 6.90]),
        "Stats.Code"
    ] = "TRADSLX!"
    te.loc[
        (te["M"] == "RCH PRG")
        & te["G"].isin(["Unknown Cancellation","Cancel At Door","Cancelled Due To Destination Location Being Closed"])
        & te["K"].isin([3.75, 3.85, 6.90]),
        "Stats.Code"
    ] = "TRADSLX!"
    te.loc[(te["M"] == "RCH ADS") & (te["O"] == "Canc. N/C") & (te["K"] == 0), "Stats.Code"] = "TRADSNC!"
    te.loc[
        (te["M"] == "RCH ADS")
        & te["G"].isin(["Advance Cancellation","Cancelled Due To Destination Location Being Closed"]),
        "Stats.Code"
    ] = "TRADSNC!"
    te.loc[(te["M"] == "RCH ADS") & (te["O"] == "FareWaived") & (te["K"] == 0), "Stats.Code"] = "TRADSFW!"

    # Rest
    te.loc[(te["M"] == "NORMAL") & (te["G"] == "Advance Cancellation"), "Stats.Code"] = "TRDNC!"

    # 1–5
    base_norm = te["M"] == "NORMAL"
    g_perf_sched = te["G"].isin(["Performed", "Scheduled"])
    g_psu = te["G"].isin(["Performed", "Unknown", "Arrived"])
    q_le_5 = te["Q"] <= 5

    te.loc[base_norm & g_perf_sched & q_le_5 & (te["K"] == 3.85), "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_perf_sched & q_le_5 & (te["K"] == 6.6),  "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_perf_sched & q_le_5 & (te["K"] == 4.35), "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_perf_sched & q_le_5 & (te["K"] == 6.75), "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_psu        & q_le_5 & (te["K"] == 6.6),  "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_psu        & q_le_5 & (te["K"] == 6.75), "Stats.Code"] = "TRD1!"
    te.loc[base_norm & te["G"].isin(["Performed","Unknown","Arrived","Scheduled"]) & q_le_5 & (te["K"] == 6.90), "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_psu        & q_le_5 & (te["K"] == 11),   "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_psu        & q_le_5 & (te["K"] == 13.8), "Stats.Code"] = "TRD1!"
    te.loc[base_norm & g_perf_sched & q_le_5 & (te["K"] == 4),    "Stats.Code"] = "TRADS!"

    g_uncancel = te["G"].isin(["Unknown Cancellation", "Cancel At Door"])
    g_uncancel_plus = te["G"].isin(["Unknown Cancellation", "Cancel At Door", "Cancelled Due To Destination Location Being Closed"])

    te.loc[base_norm & g_uncancel      & q_le_5 & (te["K"] == 6.6),  "Stats.Code"] = "TRD1LX!"
    te.loc[base_norm & g_uncancel      & q_le_5 & (te["K"] == 13.8), "Stats.Code"] = "TRD1LX!"
    te.loc[base_norm & g_uncancel_plus & q_le_5 & (te["K"] == 6.75), "Stats.Code"] = "TRD1LX!"
    te.loc[base_norm & g_uncancel_plus & q_le_5 & (te["K"] == 6.90), "Stats.Code"] = "TRD1LX!"
    te.loc[base_norm & g_uncancel      & q_le_5 & (te["K"] == 13.5), "Stats.Code"] = "TRD1LX!"
    te.loc[base_norm & g_uncancel      & q_le_5 & (te["K"] == 4),    "Stats.Code"] = "TRD1LX!"
    te.loc[base_norm & g_uncancel      & q_le_5 & (te["K"] == 3.6),  "Stats.Code"] = "TRD1LX!"
    te.loc[base_norm & g_uncancel      & q_le_5 & (te["K"] == 3.85), "Stats.Code"] = "TRD1LX!"

    # 5–10
    q_5_10 = (te["Q"] > 5) & (te["Q"] <= 10)
    g_psa = te["G"].isin(["Performed", "Scheduled", "Arrived"])

    te.loc[base_norm & g_psa        & q_5_10 & (te["K"] == 10.5),  "Stats.Code"] = "TRD2!"
    te.loc[base_norm & g_psa        & q_5_10 & (te["K"] == 10.75), "Stats.Code"] = "TRD2!"
    te.loc[base_norm & g_psa        & q_5_10 & (te["K"] == 11),    "Stats.Code"] = "TRD2!"
    te.loc[base_norm & g_perf_sched & q_5_10 & (te["K"] == 0),     "Stats.Code"] = "TRD2!"
    te.loc[base_norm & g_perf_sched & q_5_10 & (te["K"] == 21.5),  "Stats.Code"] = "TRD2!"
    te.loc[base_norm & g_perf_sched & q_5_10 & (te["K"] == 22),    "Stats.Code"] = "TRD2!"

    g_uncancel_or_unknown = te["G"].isin(["Unknown Cancellation", "Cancel At Door", "Unknown"])
    te.loc[base_norm & g_uncancel_or_unknown & q_5_10 & (te["K"] == 10.5),  "Stats.Code"] = "TRD2LX!"
    te.loc[base_norm & te["G"].isin(["Unknown Cancellation","Cancel At Door","Unknown","Cancelled Due To Destination Location Being Closed"]) & q_5_10 & (te["K"] == 10.75), "Stats.Code"] = "TRD2LX!"
    te.loc[base_norm & te["G"].isin(["Unknown Cancellation","Cancel At Door","Unknown","Cancelled Due To Destination Location Being Closed"]) & q_5_10 & (te["K"] == 11),    "Stats.Code"] = "TRD2LX!"
    te.loc[base_norm & g_uncancel & q_5_10 & (te["K"] == 0), "Stats.Code"] = "TRD2LX!"

    # 10–15
    q_le_5_again = te["Q"] <= 5
    q_10_15 = (te["Q"] > 10) & (te["Q"] <= 15)
    te.loc[base_norm & g_perf_sched & q_le_5_again & (te["K"] == 13.2),  "Stats.Code"] = "TRD3!"
    te.loc[base_norm & g_perf_sched & q_le_5_again & (te["K"] == 13.5),  "Stats.Code"] = "TRD3!"
    te.loc[base_norm & g_perf_sched & q_10_15      & (te["K"] == 15.25), "Stats.Code"] = "TRD3!"
    te.loc[base_norm & g_perf_sched & q_10_15      & (te["K"] == 15.6),  "Stats.Code"] = "TRD3!"
    te.loc[base_norm & g_perf_sched & q_10_15      & (te["K"] == 15.95), "Stats.Code"] = "TRD3!"

    te.loc[base_norm & g_uncancel_or_unknown & q_10_15 & (te["K"] == 15.25), "Stats.Code"] = "TRD3LX!"
    te.loc[base_norm & g_uncancel_or_unknown & q_10_15 & (te["K"] == 15.6),  "Stats.Code"] = "TRD3LX!"
    te.loc[base_norm & g_uncancel_or_unknown & q_10_15 & (te["K"] == 15.95), "Stats.Code"] = "TRD3LX!"

    # 15–20
    q_15_20 = (te["Q"] > 15) & (te["Q"] <= 20)
    te.loc[base_norm & g_perf_sched & q_15_20 & (te["K"] == 18.75), "Stats.Code"] = "TRD4!"
    te.loc[base_norm & g_perf_sched & q_15_20 & (te["K"] == 19.65), "Stats.Code"] = "TRD4!"

    te.loc[base_norm & g_uncancel & q_15_20 & (te["K"] == 18.75), "Stats.Code"] = "TRD4LX!"
    te.loc[base_norm & g_uncancel & q_15_20 & (te["K"] == 19.25), "Stats.Code"] = "TRD4LX!"
    te.loc[base_norm & g_uncancel & q_15_20 & (te["K"] == 19.65), "Stats.Code"] = "TRD4LX!"

    # 20–25
    q_20_25 = (te["Q"] > 20) & (te["Q"] <= 25)
    te.loc[base_norm & g_perf_sched & q_20_25 & (te["K"] == 23), "Stats.Code"] = "TRD5!"
    te.loc[base_norm & g_uncancel   & q_20_25 & (te["K"] == 23), "Stats.Code"] = "TRD5LX!"

    # 25<
    q_gt_25 = te["Q"] > 25
    te.loc[base_norm & g_perf_sched & q_gt_25, "Stats.Code"] = "TRD6!"
    te.loc[base_norm & g_uncancel   & q_20_25, "Stats.Code"] = "TRD6LX!"
    te.loc[base_norm & g_uncancel   & q_gt_25, "Stats.Code"] = "TRD6LX!"

    # HAL
    te.loc[te["K"] == 26.27, "Stats.Code"] = "TRHAL"

    # EHC / Shopping
    te.loc[(te["K"] == 3.6)  & te["G"].isin(["Performed", "Scheduled"]), "Stats.Code"] = "TRPRG!"
    te.loc[(te["K"] == 3.7)  & te["G"].isin(["Performed", "Scheduled"]), "Stats.Code"] = "TRPRG!"
    te.loc[(te["K"] == 3.85) & te["G"].isin(["Performed", "Scheduled"]), "Stats.Code"] = "TRPRG!"

    te.loc[(te["K"] == 0.0) & te["G"].isin(["Unknown Cancellation","Cancel At Door","Performed"]), "Stats.Code"] = "TRPRGLX!"
    te.loc[(te["K"] == 3.6) & te["G"].isin(["Unknown Cancellation","Cancel At Door"]), "Stats.Code"] = "TRPRGLX!"
    te.loc[(te["K"] == 3.7) & te["G"].isin(["Unknown Cancellation","Cancel At Door"]), "Stats.Code"] = "TRPRGLX!"

    te.loc[(te["M"] == "NORMAL") & (te["O"] == "FareWaived"), "Stats.Code"] = "TRPRGFW!"
    te.loc[(te["M"] == "NORMAL") & (te["O"] == "Canc. N/C"),  "Stats.Code"] = "TRPRGNC!"
    te.loc[(te["M"] == "NORMAL") & (te["O"] == "No Charge"),  "Stats.Code"] = "TRPRGNC!"

    # drop column 14 (R 1-based) -> Python index 13
    te = te.drop(columns=[te.columns[13]])

    # reorder: c(1:10,12,13,11,14:25) -> [:10], 11,12,10, 13:25
    te = te[list(te.columns[:10]) + [te.columns[11], te.columns[12], te.columns[10]] + list(te.columns[13:25])]

    # D: prepend zeros, keep rightmost 7
    te["D"] = ("000000" + te["D"].astype(str)).str[-7:]

    # reorder: c(1,2,4,3,5:25) -> 0,1,3,2,4..24
    cols = list(te.columns)
    te = te[[cols[0], cols[1], cols[3], cols[2], *cols[4:25]]].copy()

    # rename 3rd/4th to C, D
    cols = list(te.columns)
    cols[2] = "C"
    cols[3] = "D"
    te.columns = cols

    # set K = 0 when N == "ESCORT"
    te.loc[te["N"] == "ESCORT", "K"] = 0

    # drop codes + ESCORT rows
    drop_codes = [
        "TRADSFW!", "TRADSNC!", "TRDFW!", "TRDNC!", "TRPRGFW!",
        "TRPRGNC!", "TRSHPFW!", "TRSHPNC!", "TRADSFW!",
    ]
    te = te[~te["Stats.Code"].isin(drop_codes)]
    te = te[te["N"] != "ESCORT"]

    # fix E
    te["E"] = te["E"].astype(str).replace({
        "000669627": "AC000669627",
        "000005592": "AC000005592",
        "000007326": "AC000007326",
        "000005078": "AC000005078",
        "000007449": "AC000007449",
        "000006174": "AC000006174",
        "000007538": "AC000007538",
        "000006557": "1017294",
        "AC000006557": "1017294",
        "001096265": "AC001096265",
        "001345": "0013445",
    })

    # ===== AlayaCare ADT join logic =====
    ALAYA_DIR_TPL = r"C:\Users\=\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\AlayaCare"
    ALAYA_DIR = expand_user_placeholder(ALAYA_DIR_TPL, usern)

    adt = pd.read_csv(ALAYA_DIR / "ADT with Zone and Reference.csv", dtype="object")
    adt["ProcuraID"] = np.where(adt["ClientID"] == "AC000945993", "1014989", adt["ProcuraID"])
    adt = adt[adt["ServiceCodeDept"] == "Transportation"].copy()
    adt = adt.iloc[:, [2, 24, 26, 25]].copy()  # 3,25,27,26 (R) -> 2,24,26,25
    adt.columns = ["AID", "PID", "NewID", "Client"]

    # robust blank detection for PID
    pid_blank = adt["PID"].isna() | (adt["PID"].astype(str).str.strip().replace({"nan": ""}) == "")
    adt["NewPID"] = np.where(pid_blank, adt["NewID"], adt["PID"])
    adt["NewPID"] = "90000" + adt["NewPID"].astype(str)
    adt = adt.drop(columns=["NewID", "Client"]).drop_duplicates()

    atrapeze = te[te["E"].astype(str).str.startswith("AC")].copy().rename(columns={"E": "AID"})
    ptrapeze = te[~te["E"].astype(str).str.match(r"^AC")].copy().rename(columns={"E": "PID"})

    atrapeze = atrapeze.merge(adt, on="AID", how="left")
    ptrapeze = ptrapeze.merge(adt, on="PID", how="left")

    te = pd.concat([ptrapeze, atrapeze], ignore_index=True)

    # te[, c(1:4, 27, 26, 5:25)]
    cols = list(te.columns)
    te = te[[*cols[0:4], cols[26], cols[25], *cols[4:25]]].copy()

    # rename 5th to "E"
    cols = list(te.columns)
    cols[4] = "E"
    te.columns = cols

    # C <- E
    te["C"] = te["E"]

    # billable vs. not billable
    billablte = te.copy()
    billablte = billablte[billablte["E"].notna()].copy()
    notbillablete = te[te["E"].isna()].copy()  # kept for parity, not used further

    # I <- Y, drop Y/AID/PID, then set I="9"
    if "Y" in billablte.columns:
        billablte["I"] = billablte["Y"]
        billablte = billablte.drop(columns=["Y"])
    for c in ["AID", "PID"]:
        if c in billablte.columns:
            billablte = billablte.drop(columns=[c])
    billablte["I"] = "9"

    # reorder c(1:3,5,4,6:24)
    cols = list(billablte.columns)
    order = [*cols[0:3], cols[4], cols[3], *cols[5:24]]
    billablte = billablte[order].copy()

    # drop rows where 8th column ∈ {503, 500, 504, 509}
    col8 = billablte.columns[7]
    billablte = billablte[~pd.to_numeric(billablte[col8], errors="coerce").isin([503, 500, 504, 509])].copy()

    # ===== Output to Downloads =====
    downloads = Path.home() / "Downloads"
    downloads.mkdir(parents=True, exist_ok=True)
    out_path = downloads / "trapezerides.csv"

    billablte.fillna("NA").to_csv(out_path, sep=",", index=False, header=False)

    if excel_sheet:
        print(f"Success! Loaded: {latest_file.name} | sheet: {excel_sheet} → wrote: {out_path}")
    else:
        print(f"Success! Loaded: {latest_file.name} (CSV) → wrote: {out_path}")

    return out_path


def main():
    try:
        run()
    except Exception as e:
        print(f"FAILED: {e}")
        raise


if __name__ == "__main__":
    main()
