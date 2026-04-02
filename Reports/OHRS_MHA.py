# ohrs_mha_blocks.py
from __future__ import annotations

import math
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Sequence, Tuple

import pandas as pd
import numpy as np
import xlsxwriter  # imported to mirror your style; not used for CSV

# =========================
#        CONSTANTS
# =========================
ORG_ROOT = fr"C:\Users\{os.getlogin()}\OneDrive - Reconnect Community Health Services (1)"
TREAT_CLEAN     = Path(ORG_ROOT) / r"data\Clean Data\Treat"
TARGETS_CSV     = Path(ORG_ROOT) / r"data\Templates\OHRS_MHA targets.csv"
OUTPUT_CSV      = Path(ORG_ROOT) / r"data\Clean Data\OHRS\OHRS_MHA.csv"

FY_START_MONTH = 4  # April

# =========================
#     GENERIC HELPERS
# =========================
def read_csv(path: Path, **kwargs) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing file: {path}")
    defaults = dict(encoding="utf-8-sig", keep_default_na=False)
    defaults.update(kwargs)
    return pd.read_csv(path, **defaults)

def coerce_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")

def _fmt_cell(v) -> str:
    """Format body/total/target cells: ints; blanks as '-'."""
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return "-"
        f = float(v)
        if abs(f) < 0.5:
            return "-"
        return f"{int(round(f))}"
    except Exception:
        return "-"

def percent_remaining_signed(actual: Optional[float], target: Optional[float]) -> Optional[str]:
    """(Target - Actual)/Target as %-string. Allows negative when over target."""
    try:
        if target is None:
            return ""
        t = float(target)
        if t == 0 or actual is None:
            return ""
        a = float(actual)
        rem = (t - a) / t
        return f"{rem*100:.2f}%"
    except Exception:
        return ""

# =========================
#     TARGETS LOADER
# =========================
def _parse_targets_csv(targets_path: Path) -> pd.DataFrame | None:
    if not targets_path.exists():
        print(f"[WARN] Targets file not found at {targets_path}")
        return None

    raw = pd.read_csv(targets_path, header=None, dtype=str, keep_default_na=False)
    if raw.empty:
        print(f"[WARN] Targets file is empty: {targets_path}")
        return None

    alias_to_canon = {
        "Service Provider Interactions": "Service Provider Interactions",
        "SPI":  "Service Provider Interactions",
        "SPIs": "Service Provider Interactions",
        "Non-Unique SPIs": "Service Provider Interactions",

        "Visits": "Visits",
        "Attendance Days": "Attendance Days",
        "Resident Days": "Resident Days",

        "Individuals Served": "Individuals Served",

        "Service Provider Group Interactions": "Service Provider Group Interactions",
        "SPGI":   "Service Provider Group Interactions",
        "SPIGIs": "Service Provider Group Interactions",

        "Group Session": "Group Session",
        "Group Participants": "Group Participants",
    }

    canon_order = [
        "Functional Center", "Label",
        "Service Provider Interactions", "Visits", "Attendance Days", "Resident Days",
        "Individuals Served", "Service Provider Group Interactions", "Group Session", "Group Participants",
    ]

    def _coerce_num(x: str):
        x = (x or "").strip()
        if x in ("", "-", "–", "—"):
            return pd.NA
        try:
            x_clean = x.replace(",", "")
            f = float(x_clean)
            i = int(f)
            return i if f.is_integer() else f
        except Exception:
            return pd.NA

    recs = []
    i, n = 0, len(raw)
    while i + 1 < n:
        header_row = raw.iloc[i].fillna("")
        value_row  = raw.iloc[i+1].fillna("")

        fc = str(header_row.iloc[0]).strip()
        label = str(value_row.iloc[0]).strip()

        hdr_cols = [str(c).strip() for c in header_row.iloc[2:].tolist() if str(c).strip() != ""]
        vals     = value_row.iloc[2:2+len(hdr_cols)].tolist()

        rec = {"Functional Center": fc, "Label": label}
        for raw_name, v in zip(hdr_cols, vals):
            canon = alias_to_canon.get(raw_name, raw_name)
            rec[canon] = _coerce_num(str(v))
        recs.append(rec)
        i += 2

    targets = pd.DataFrame.from_records(recs)
    for c in canon_order:
        if c not in targets.columns:
            targets[c] = pd.NA
    targets = targets[canon_order]
    print(f"[TARGETS] Parsed {len(targets)} rows from {targets_path.name}")
    return targets

# =========================
#       CSV "WRITER"
# =========================
class SheetWriter:
    """
    CSV-backed writer that lays out blocks:
    - Title+headers row (col0=title, col1..N=headers)
    - Program rows with numbers
    - YTD total, FY target, Remaining
    - Blank spacer
    """
    def __init__(self, out_path: Path):
        self.out_path = out_path
        self.table: list[list[Optional[str]]] = []
        self.row = 0

    def _ensure_row(self, r: int, cols: int):
        while len(self.table) <= r:
            self.table.append([])
        cur = len(self.table[r])
        if cur < cols:
            self.table[r].extend([None] * (cols - cur))

    def _normalize(self):
        max_cols = max((len(r) for r in self.table), default=0)
        for r in self.table:
            if len(r) < max_cols:
                r.extend([None] * (max_cols - len(r)))
        for r in range(len(self.table)):
            self.table[r] = [("" if v is None else v) for v in self.table[r]]

    def finalize(self) -> pd.DataFrame:
        self._normalize()
        df = pd.DataFrame(self.table)
        self.out_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(self.out_path, index=False, header=False, encoding="utf-8-sig")
        print(f"[SAVED] {self.out_path}")
        return df

    def write_section(self, title: str, headers: Sequence[str], row_labels: Sequence[str]) -> int:
        cols = 1 + len(headers)
        self._ensure_row(self.row, cols)
        self.table[self.row][0] = title
        for j, h in enumerate(headers, start=1):
            self.table[self.row][j] = h
        self.row += 1

        start = self.row
        for lab in row_labels:
            self._ensure_row(self.row, cols)
            self.table[self.row][0] = lab
            self.row += 1

        self._ensure_row(self.row, 1)  # spacer row
        self.row += 1
        return start

    def write_val(self, r: int, c: int, val: Optional[str | float]):
        if val is None:
            return
        self._ensure_row(r, c + 1)
        self.table[r][c] = val

# =========================
#      BUILD THE TABLES
# =========================
# Canonical → display headers
DEFAULT_HEADERS = {
    "Service Provider Interactions": "SPI",
    "Visits": "Visits",
    "Attendance Days": "Attendance Days",
    "Resident Days": "Resident Days",
    "Individuals Served": "Individuals Served",
    "Service Provider Group Interactions": "SPIGIs",
    "Group Session": "Group Sessions",
    "Group Participants": "Group Participants",
}

def _headers_for_fc(fc: str) -> dict[str, str]:
    if fc == "725517620 - Family Initiatives":
        h = DEFAULT_HEADERS.copy()
        h["Service Provider Interactions"] = "Non-Unique SPIs"
        return h
    return DEFAULT_HEADERS

# Allowed metrics per FC
ORDER_DEFAULT = [
    "Service Provider Interactions", "Visits", "Individuals Served",
    "Service Provider Group Interactions", "Group Session", "Group Participants"
]
ORDER_SAFE_BEDS = ["Service Provider Interactions", "Resident Days", "Individuals Served"]
ORDER_FAMILY = ["Service Provider Interactions", "Attendance Days", "Individuals Served"]

def _allowed_order_for_fc(fc: str) -> list[str]:
    if fc == "725407660 - Safe Beds":
        return ORDER_SAFE_BEDS
    if fc == "725517620 - Family Initiatives":
        return ORDER_FAMILY
    if fc == "725097811 - Addictions":
        return ORDER_DEFAULT
    return ORDER_DEFAULT

# Where to find targets for each metric inside the targets CSV (aliases)
def _target_alias_map_for_fc(fc: str) -> dict[str, list[str]]:
    base = {
        "Service Provider Interactions": ["Service Provider Interactions"],
        "Visits": ["Visits"],
        "Attendance Days": ["Attendance Days"],
        "Resident Days": ["Resident Days"],
        "Individuals Served": ["Individuals Served"],
        "Service Provider Group Interactions": ["Service Provider Group Interactions", "SPIGIs", "SPGI"],
        "Group Session": ["Group Session"],
        "Group Participants": ["Group Participants"],
    }
    # Addictions: allow Visits target to fall back to Attendance Days (AD is not shown, but used if Visit target missing)
    if fc == "725097811 - Addictions":
        base["Visits"] = ["Visits", "Attendance Days"]
    return base

def _pick_target_for_metric(r: pd.Series, metric: str, am: dict[str, list[str]]) -> float | None:
    for cand in am.get(metric, [metric]):
        if cand in r.index:
            v = pd.to_numeric(r.get(cand), errors="coerce")
            if pd.notna(v):
                return float(v)
    return None

def build_fc_blocks(
    main_table: pd.DataFrame,
    targets_norm: pd.DataFrame,
    out_csv: Path,
    fc_totals_override: dict[str, dict[str, float]] | None = None,  # NEW (optional)
) -> pd.DataFrame:
    """
    Build OHRS blocks per Functional Center and write to CSV.

    fc_totals_override (optional):
        A nested dict allowing FC-level YTD overrides per metric.
        Shape: { "<MetricName>": { "<Functional Center>": <float YTD value>, ... }, ... }
        Example for de-duped Individuals Served:
            {
              "Individuals Served": {
                "7250976 - Case Management": 812.0,
                "725107620 - ACTT": 423.0,
                ...
              }
            }
    """
    sw = SheetWriter(out_csv)

    # targets lookup
    t_lookup = (
        targets_norm.set_index("Functional Center", drop=False)
        if targets_norm is not None and not targets_norm.empty
        else None
    )

    # process FCs in appearance order
    for fc, g in main_table.groupby("Functional Center", sort=False):
        headers_map = _headers_for_fc(fc)
        order = _allowed_order_for_fc(fc)

        # Metrics present in data & allowed for FC
        present = [m for m in order if m in g.columns]

        # Keep metrics that either have any nonzero YTD or have a target
        keep: list[str] = []
        r_target = t_lookup.loc[fc] if (t_lookup is not None and fc in t_lookup.index) else None
        am = _target_alias_map_for_fc(fc)

        for m in present:
            body_sum = pd.to_numeric(g[m], errors="coerce").sum()
            has_target = False
            if r_target is not None:
                tv = _pick_target_for_metric(r_target, m, am)
                has_target = (tv is not None)
            if (body_sum > 0) or has_target:
                keep.append(m)

        # If nothing to keep, just title + spacer
        if not keep:
            sw.write_section(fc, [], [])
            continue

        # Display headers
        display_headers = [headers_map[m] for m in keep]

        # Program labels – drop programs that are all zero (across kept metrics)
        g2 = g[["Program"] + keep].copy()
        keep_mask = g2[keep].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) > 0
        body = g2.loc[keep_mask].reset_index(drop=True)

        # Row labels = programs + totals + target + remaining
        row_labels = (
            body["Program"].tolist()
            + ["Year to Date total", f"{fc.split(' - ', 1)[-1]} Fiscal Year Target", f"{fc.split(' - ', 1)[-1]} Remaining"]
        )
        start_r = sw.write_section(fc, display_headers, row_labels)

        # 1) Body rows
        for i, (_, row) in enumerate(body.iterrows()):
            for j, m in enumerate(keep, start=1):
                sw.write_val(start_r + i, j, _fmt_cell(row[m]))

        # 2) YTD totals row (wrap numbers in square brackets)
        ytd_row_idx = start_r + len(body)
        for j, m in enumerate(keep, start=1):
            # default: sum of program values
            ytd_val = pd.to_numeric(g[m], errors="coerce").sum()

            # OVERRIDE for FC-level totals (e.g., Individuals Served)
            if fc_totals_override and m in fc_totals_override and fc in fc_totals_override[m]:
                ytd_val = float(fc_totals_override[m][fc])

            vtxt = _fmt_cell(ytd_val)
            if vtxt != "-":
                vtxt = f"[{vtxt}]"
            sw.write_val(ytd_row_idx, j, vtxt)

        # 3) Target & 4) Remaining rows
        tgt_row_idx = ytd_row_idx + 1
        rem_row_idx = ytd_row_idx + 2

        # Build totals used for Remaining% (must match the YTD row value)
        totals_by_metric: dict[str, float] = {}
        for m in keep:
            total_val = pd.to_numeric(g[m], errors="coerce").sum()
            if fc_totals_override and m in fc_totals_override and fc in fc_totals_override[m]:
                total_val = float(fc_totals_override[m][fc])
            totals_by_metric[m] = total_val

        for j, m in enumerate(keep, start=1):
            tv = _pick_target_for_metric(r_target, m, am) if r_target is not None else None

            # Wrap Fiscal Year Target values in { }  (keep "-" as-is)
            vtxt = _fmt_cell(tv)
            if vtxt != "-":
                vtxt = f"{{{vtxt}}}"   # double braces give literal { } in f-strings
            sw.write_val(tgt_row_idx, j, vtxt)

            # Remaining uses totals_by_metric (already override-aware)
            sw.write_val(rem_row_idx, j, percent_remaining_signed(totals_by_metric[m], tv))

    return sw.finalize()


# =========================
#            RUN
# =========================
def run(output_path: Optional[Path | str] = None) -> pd.DataFrame:
    out_path = Path(output_path) if output_path is not None else OUTPUT_CSV

    # --- Load MIS (program totals by month) ---
    mis_files = sorted(TREAT_CLEAN.glob("rpt_MIS_Stats*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not mis_files:
        raise FileNotFoundError(f"No 'rpt_MIS_Stats*.csv' found in {TREAT_CLEAN}")
    mis_path = mis_files[0]
    print(f"[MIS] Using: {mis_path.name}")

    mis_df = read_csv(mis_path)
    mis_df = mis_df[mis_df['Category'].isin([
        'Service Provider Interactions',
        'Service Provider Group Interactions',
        'Attendance Days Face to Face',
        'Attendance Days Non Face to Face',
        'Resident Days',
        'Group Participants',
        'Group Session',
        'Individuals Serviced',
        'Individuals Served by Org'
    ])]

    mis_df2 = mis_df.drop(columns=["FCC", "MIS Code", "Description"], errors="ignore").copy()

    FY_MONTHS = ["Apr","May","June","July","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
    month_cols = [c for c in FY_MONTHS if c in mis_df2.columns]
    mis_df2[month_cols] = mis_df2[month_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    combined = mis_df2.groupby(["Category", "Program"], as_index=False)[month_cols].sum()

    today = pd.Timestamp.today().normalize()
    current_lbl = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"June",7:"July",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}[today.month]
    cutoff_idx = FY_MONTHS.index(current_lbl)
    months_to_sum = [m for m in FY_MONTHS[:cutoff_idx+1] if m in combined.columns]
    combined["Total until today"] = combined[months_to_sum].sum(axis=1)

    tbl = combined[["Category", "Program", "Total until today"]].copy()

    pivot = (
        tbl.pivot_table(
            index="Program",
            columns="Category",
            values="Total until today",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()
    )
    pivot.columns.name = None

    # Attendance Days = F2F + NF2F
    pivot["Attendance Days"] = (
        pivot.filter(items=["Attendance Days Face to Face"]).sum(axis=1) +
        pivot.filter(items=["Attendance Days Non Face to Face"]).sum(axis=1)
    )
    pivot = pivot.drop(columns=["Attendance Days Face to Face","Attendance Days Non Face to Face"], errors="ignore")

    # -------------------------
# Canonicalize MIS columns
# -------------------------

# 1) Individuals Served: prefer "Individuals Serviced", else "Individuals Served by Org"
    if "Individuals Serviced" in pivot.columns:
        pivot["Individuals Served"] = pd.to_numeric(pivot["Individuals Serviced"], errors="coerce").fillna(0)
    elif "Individuals Served by Org" in pivot.columns:
        pivot["Individuals Served"] = pd.to_numeric(pivot["Individuals Served by Org"], errors="coerce").fillna(0)
    else:
        pivot["Individuals Served"] = 0

    # Drop the raw MIS Individuals columns to avoid accidental use later
    pivot = pivot.drop(columns=[c for c in ["Individuals Serviced", "Individuals Served by Org"] if c in pivot.columns])

    # 2) (Optional) Rename any other known aliases WITHOUT summing
    #    (Nothing else needs aliasing here—MIS uses canonical names for other metrics—
    #     but leaving an example if you ever add an alias.)
    other_renames = {
        # "Some Alias": "Canonical Name",
    }
    for src, dst in other_renames.items():
        if src in pivot.columns and src != dst:
            if dst not in pivot.columns:
                pivot = pivot.rename(columns={src: dst})
            else:
                # If both somehow exist, keep dst and drop src (no adding)
                pivot = pivot.drop(columns=[src])


    # =========================
    # Add Visits (from rpt_ProgStaffFinance) — FY window & all logic based on Interaction Start Date
    # =========================
    cands = sorted(TREAT_CLEAN.glob("rpt_ProgStaffFinance.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not cands:
        raise FileNotFoundError("No 'rpt_ProgStaffFinance.csv' found in Clean Data\\Treat")
    df_progress = read_csv(cands[0], dtype={"ID": "string"})

    # Fiscal year window (Interaction Start Date)
    int_dt = pd.to_datetime(df_progress["Interaction Start Date"], errors="coerce")
    apr1_this_year = pd.Timestamp(today.year, 4, 1)
    fy_start = apr1_this_year if today >= apr1_this_year else pd.Timestamp(today.year - 1, 4, 1)
    mask = (int_dt >= fy_start) & (int_dt <= today)

    # Slice for Individuals Served (unique per FC) and Visits (de-dupe by type/day)
    df_clients_fy = df_progress.loc[mask, ["ClientName", "Program"]].copy()
    df_progress_fy = df_progress.loc[mask].copy()

    # --- Visits: one per (Client, Interaction-Day, Visit Type), across all programs ---
    df_progress_fy["Interaction Start Date"] = pd.to_datetime(df_progress_fy["Interaction Start Date"], errors="coerce")
    df_progress_fy["_day"] = df_progress_fy["Interaction Start Date"].dt.normalize()

    df_progress_fy["ClientName"] = df_progress_fy["ClientName"].astype(str).str.strip()
    df_progress_fy["Visit Type"] = df_progress_fy["Visit Type"].astype(str).str.strip().replace({"": "Unknown"})

    df_progress_fy = (
        df_progress_fy
        .drop_duplicates(subset=["ClientName", "_day", "Visit Type"], keep="first")
        .reset_index(drop=True)
    )

    # Visits per Program
    df_prog_visits = (
        df_progress_fy.groupby("Program", dropna=False).size()
        .reset_index(name="Visits")
    )

    # Merge visits into main_table
    main_table = pivot.merge(df_prog_visits, how="left", on="Program").fillna(0)

    # Ensure ints
    for c in [
        "Service Provider Interactions","Service Provider Group Interactions",
        "Group Session","Group Participants","Resident Days",
        "Individuals Served","Attendance Days","Visits"
    ]:
        if c in main_table.columns:
            main_table[c] = pd.to_numeric(main_table[c], errors="coerce").fillna(0).astype(int)

    # ---- Functional Center mapping ----
    fc_ACTT = [
        "ACT Case Management","ACT Social Group","ACTT","CAMH at Home","CBTP",
        "CDBT Skills Group","DBT","DBT 101 Skills Group","DBT Groups","Harborlight",
        "Mindfulness Group","PG ACTT","PGACTT Case Management","Psychiatry Step Down","SoTo","SOTO Social Group",
        "WeTo","WeTo Social Group"
    ]
    fc_CM = [
        "Addiction Focused P-E Group","Addiction Support Services","Art Therapy",
        "Case Management","Clinical Nutrition Serivces","Clinical Nutrition Services",
        "Continuing Services","Continuing Service","CWCD",
        "Dog Assisted Therapy Group","Dog assisted Therapy Group","Drama Therapy",
        "HASP CM","Mainstay","MHJ CM","Mild to Moderate CM","Music Therapy",
        "Primary Care InterProfessional","RASS","SHEY CM","Transitional Aged Youth",
        "UHN Geriatric Psychiatry Clini","Virtual Case Management"
    ]
    fc_Forensic = ["Forensic CM","NFTCM"]
    fc_Crisis   = ["COSS Follow Up","COSS Crisis","FOCUS","JAM","Prevention CM"]
    fc_FAME     = ["FAME","FAME Group","FAMEKids","FAMEKids Group"]
    fc_STCB     = ["STCB"]                 # Safe Beds
    fc_COPA     = ["COPA","ASH/COPA","ASH/COPA Skill Group","COPA SMART Recovery Group"]  # Addictions

    fc_map: dict[str, str] = {}
    fc_map.update({p: "725107620 - ACTT" for p in fc_ACTT})
    fc_map.update({p: "7250976 - Case Management" for p in fc_CM})
    fc_map.update({p: "725107655 - Forensic CM" for p in fc_Forensic})
    fc_map.update({p: "7251576 - Crisis" for p in fc_Crisis})
    fc_map.update({p: "725407660 - Safe Beds" for p in fc_STCB})
    fc_map.update({p: "725097811 - Addictions" for p in fc_COPA})
    fc_map.update({p: "725517620 - Family Initiatives" for p in fc_FAME})

    # Map Program -> FC for main_table
    main_table["Program"] = main_table["Program"].astype(str).str.strip()
    main_table["Functional Center"] = main_table["Program"].map(lambda x: fc_map.get(x, "")).fillna("")
    main_table = main_table[main_table["Functional Center"] != ""].reset_index(drop=True)


    # Targets
    targets_norm = _parse_targets_csv(TARGETS_CSV)

    df_out = build_fc_blocks(main_table, targets_norm, OUTPUT_CSV)

    return df_out



# =========================
#           MAIN
# =========================
def main() -> None:
    _ = run()

if __name__ == "__main__":
    main()
