# ohrs_css_pipeline.py
from __future__ import annotations

import math
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Sequence, Tuple

import pandas as pd
import xlsxwriter  # kept; not used by CSV writer but left to avoid broader edits


# =========================
#        CONSTANTS
# =========================
ORG_ROOT = fr"C:\Users\{os.getlogin()}\OneDrive - Reconnect Community Health Services (1)"
ALAYACARE_CLEAN = Path(ORG_ROOT) / r"data\Clean Data\AlayaCare"
TREAT_CLEAN     = Path(ORG_ROOT) / r"data\Clean Data\Treat"
TARGETS_CSV     = Path(ORG_ROOT) / r"data\Templates\OHRS_SFS targets.csv"
OUTPUT_XLSX     = Path(fr"C:\Users\{os.getlogin()}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\OHRS\OHRS_SFS.csv")

FY_START_MONTH = 4  # April


# =========================
#     GENERIC HELPERS
# =========================
def read_csv(path: Path, **kwargs) -> pd.DataFrame:
    """Robust CSV reader."""
    if not path.exists():
        raise FileNotFoundError(f"Missing file: {path}")
    defaults = dict(encoding="utf-8-sig")
    defaults.update(kwargs)
    return pd.read_csv(path, **defaults)


def fiscal_window_until_last_month(today: Optional[pd.Timestamp] = None) -> Tuple[pd.Timestamp, pd.Timestamp]:
    """Return (start_date, end_date) for the fiscal year (Apr 1 – Mar 31),
    ending on today's date."""
    today = (today or pd.Timestamp.today()).normalize()
    fy_start_year = today.year if today.month >= FY_START_MONTH else today.year - 1
    start_date = pd.Timestamp(fy_start_year, FY_START_MONTH, 1)
    end_date = today
    return start_date, end_date


def coerce_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def nunique_safe(series: pd.Series) -> int:
    try:
        return int(series.nunique(dropna=True))
    except Exception:
        return 0


def sum_none(*vals) -> Optional[float]:
    """Sum numbers, ignoring None/NaN; return None if all missing."""
    nums = []
    for v in vals:
        try:
            if v is None:
                continue
            fv = float(v)
            if math.isnan(fv):
                continue
            nums.append(fv)
        except Exception:
            continue
    return sum(nums) if nums else None


def percent_remaining(actual: Optional[float], target: Optional[float]) -> Optional[float]:
    """
    (Target - Actual) / Target as a fraction (e.g., 0.40). Returns None if target is None/<=0 or actual is None.
    Floors at 0 if actual >= target.
    """
    try:
        if target is None or target <= 0 or actual is None:
            return None
        rem = (float(target) - float(actual)) / float(target)
        return max(rem, 0.0)
    except Exception:
        return None


# =========================
#     TARGETS LOADER
# =========================
def _norm_measure(label: str) -> str:
    m = (label or "").strip()
    aliases = {
        "Individual Served": "Individuals Served",
        "Individuals Served": "Individuals Served",
        "Individuals Served (Group Participant)": "Individuals Served",
        "Meal Delivered Combined": "Meal Delivered Combined",
        "Attendance Days": "Attendance Days",
        "Hours of Care": "Hours of Care",
        "Group Session": "Group Sessions",
        "Group Sessions": "Group Sessions",
        "Inpatient/Resident Days": "Inpatient/Resident Days",
        "SPGIs": "SPGIs",
        "Visits": "Visits",
    }
    return aliases.get(m, m)


def load_targets(csv_path: Path) -> Dict[str, Dict[str, Optional[float]]]:
    """
    Reads a 2-row-per-program CSV:
      Row i:    headers (Program in col 0, measures in cols 1..)
      Row i+1:  values  (ends with 'Fiscal Year Target')
    Returns: targets[program][measure] -> float or None
    """
    raw = read_csv(csv_path, header=None, dtype=str).fillna("")
    targets: Dict[str, Dict[str, Optional[float]]] = {}

    i = 0
    while i < len(raw):
        if i + 1 >= len(raw):
            break

        header = raw.iloc[i].tolist()
        values = raw.iloc[i + 1].tolist()

        program = (header[0] or "").strip()
        if not program:
            i += 1
            continue

        # Only accept rows that indicate Fiscal Year Target
        if (values[0] or "").strip().lower().endswith("fiscal year target"):
            measures = [_norm_measure(x) for x in header[1:]]
            nums: list[Optional[float]] = []
            for v in values[1:]:
                v = (v or "").strip()
                if v in {"", "-", "—"}:
                    nums.append(None)
                else:
                    try:
                        nums.append(float(v))
                    except Exception:
                        nums.append(None)
            prog_dict = targets.setdefault(program, {})
            for m, n in zip(measures, nums):
                prog_dict[m] = n

        i += 2

    return targets


# =========================
#   DOMAIN CALCULATIONS
# =========================
@dataclass
class MealsResult:
    adp_individuals: Optional[float]
    adp_meals: Optional[float]
    mow_individuals: Optional[float]
    mow_meals_weighted: Optional[float]


def compute_meals(visits: pd.DataFrame) -> MealsResult:
    # ADP (Meals)
    adp = visits[visits["Service Code Department Name"] == "Adult Day Program"].copy()
    adp_individuals = nunique_safe(adp["AlaycareID"])
    adp_pairs = adp[["AlaycareID", "Date"]].drop_duplicates()
    adp_meals = float(len(adp_pairs))

    # Meals on Wheels (label in data: 'Meals on Wheel')
    mow = visits[visits["Service Code Department Name"] == "Meals on Wheel"].copy()
    mow_individuals = nunique_safe(mow["AlaycareID"])
    mow_norm = mow[["AlaycareID", "Date", "Visit Visit ID", "Service Code Department Name", "BillCode"]].drop_duplicates()

    # Weights
    weights = {"DELFG": 7, "DELFSF": 3, "DELF5": 10, "DELFR2": 6}
    mow_norm["BillCode"] = mow_norm["BillCode"].astype(str).str.upper().str.strip()
    mow_norm["Weight"] = mow_norm["BillCode"].map(weights).fillna(1).astype(int)
    mow_meals_weighted = float(mow_norm["Weight"].sum())

    return MealsResult(
        adp_individuals=float(adp_individuals),
        adp_meals=adp_meals,
        mow_individuals=float(mow_individuals),
        mow_meals_weighted=mow_meals_weighted,
    )


# ---- NEW: Day Services (own top-level group) ----
@dataclass
class DayServicesResult:
    adp_individuals: Optional[float]
    adp_attendance: Optional[float]


def compute_day_services(visits: pd.DataFrame) -> DayServicesResult:
    """
    Adult Day Services section:
      - filter ServiceCode == 'Adult Day Services'
      - Individuals Served = unique AlaycareID
      - Attendance Days   = unique (AlaycareID, Date)
    """
    ads = visits[visits["Service Code Department Name"] == "Adult Day Program"].copy()
    ads["Date"] = coerce_datetime(ads["Date"]).dt.date
    adp_individuals = float(nunique_safe(ads["AlaycareID"]))
    adp_attendance = float(len(ads[["AlaycareID", "Date"]].drop_duplicates()))
    return DayServicesResult(adp_individuals, adp_attendance)


@dataclass
class SocialCongregateResult:
    health_wellness_individuals: Optional[float]
    health_wellness_attendance: Optional[float]
    york_west_individuals: Optional[float]
    york_west_attendance: Optional[float]


def compute_social_congregate(visits: pd.DataFrame) -> SocialCongregateResult:
    # Health & Wellness
    hw = visits[visits["Service Code Department Name"] == "Health and Wellness"].copy()
    hw["Date"] = coerce_datetime(hw["Date"]).dt.date
    hw_ind = float(nunique_safe(hw["AlaycareID"]))
    hw_att = float(len(hw[["AlaycareID", "Date"]].drop_duplicates()))

    # York West (label: 'York W')
    yw = visits[visits["Service Code Department Name"] == "York W"].copy()
    yw["Date"] = coerce_datetime(yw["Date"]).dt.date
    yw_ind = float(nunique_safe(yw["AlaycareID"]))
    yw_att = float(len(yw[["AlaycareID", "Date"]].drop_duplicates()))

    return SocialCongregateResult(hw_ind, hw_att, yw_ind, yw_att)


@dataclass
class CombinedPSHMRespiteResult:
    home_help_individuals: Optional[float]
    home_help_hours: Optional[float]
    psw_respite_individuals: Optional[float]
    psw_respite_hours: Optional[float]


def compute_combined_ps_hm_respite(visits: pd.DataFrame) -> CombinedPSHMRespiteResult:
    # Home Help
    hh = visits[visits["Service Code Department Name"] == "Home Help Worker"].copy()
    hh_ind = float(nunique_safe(hh["AlaycareID"]))
    hh_hours = float(hh[["Duration", "Visit Visit ID"]].drop_duplicates()["Duration"].sum())

    # PSW Respite (RCVIS, RPOC)
    psw_r = visits[
        (visits["Service Code Department Name"] == "Personal Support Worker")
        & (visits["ServiceCode"].isin(["RCVIS", "RPOC"]))
    ].copy()
    psw_r_ind = float(nunique_safe(psw_r["AlaycareID"]))
    psw_r_hours = float(psw_r[["Duration", "Visit Visit ID"]].drop_duplicates()["Duration"].sum())

    return CombinedPSHMRespiteResult(hh_ind, hh_hours, psw_r_ind, psw_r_hours)


@dataclass
class AssistedLivingResult:
    psw_individuals: Optional[float]
    psw_hours: Optional[float]
    resident_days: Optional[float]


def _calc_resident_days_row(row: pd.Series, today: pd.Timestamp, fy_start: pd.Timestamp) -> int:
    start = row.get("ServiceStartDate")
    status = str(row.get("Status", "")).strip().lower() if pd.notna(row.get("Status", None)) else ""
    eff = row.get("Status (Effective) Date")

    if pd.isna(start):
        return 0

    # Determine end
    if status == "active":
        end = today
    elif status == "discharged":
        end = eff
    else:
        end = eff if pd.notna(eff) else today

    # Clamp & compute
    start = max(start, fy_start)
    end = end if pd.notna(end) else today
    if end < fy_start or end < start:
        return 0

    return max((end - start).days, 0)


def compute_assisted_living(adt_df: pd.DataFrame, visits: pd.DataFrame, today: pd.Timestamp, fy_start: pd.Timestamp) -> AssistedLivingResult:
    # PSW overall
    psw = visits[visits["Service Code Department Name"] == "Personal Support Worker"].copy()
    psw_individuals = float(nunique_safe(psw["AlaycareID"]))
    psw_hours = float(psw[["Duration", "Visit Visit ID"]].drop_duplicates()["Duration"].sum())

    # Resident Days from ADT
    psw_adt = adt_df[adt_df["ServiceCodeDept"] == "Personal Support Worker"].copy()
    psw_adt["ServiceStartDate"] = coerce_datetime(psw_adt["ServiceStartDate"])
    psw_adt["Status (Effective) Date"] = coerce_datetime(psw_adt["Status (Effective) Date"])
    psw_adt["Resident days"] = psw_adt.apply(lambda r: _calc_resident_days_row(r, today, fy_start), axis=1)
    resident_days = float(psw_adt["Resident days"].sum())

    return AssistedLivingResult(psw_individuals, psw_hours, resident_days)


@dataclass
class CaseManagementResult:
    individuals: Optional[float]
    visits_count: Optional[float]


def compute_case_management(visits: pd.DataFrame) -> CaseManagementResult:
    cm = visits[visits["Service Code Department Name"] == "Case Management"].copy()
    cm_ind = float(nunique_safe(cm["AlaycareID"]))
    cm_visits = float(cm["Visit Visit ID"].nunique())
    return CaseManagementResult(cm_ind, cm_visits)


@dataclass
class CaregiverSupportTables:
    cbt_sessions: Optional[float]
    fame_sessions: Optional[float]
    famegrp_sessions: Optional[float]
    cbt_participants: Optional[float]
    fame_participants: Optional[float]
    famegrp_participants: Optional[float]
    cg_visits: Optional[float]
    cg_participants: Optional[float]


def _get_program_value(df: Optional[pd.DataFrame], program_name: str) -> Optional[float]:
    if df is None or df.empty:
        return None
    try:
        mask = df["Program"].astype(str).str.strip() == program_name
        val = df.loc[mask, "Total"]
        if val.empty:
            return None
        return float(val.iloc[0])
    except Exception:
        return None


def compute_caregiver_support_tables(
    mis_report: pd.DataFrame,
    visits: pd.DataFrame
) -> CaregiverSupportTables:
    # MIS-derived tables
    gs = mis_report[(mis_report["Category"] == "Group Session")
                    & (mis_report["Program"].isin(["FAME", "FAME Group", "CBT Adult Group", "CBT Senior Group"]))] \
                    .groupby("Program", as_index=False)["Total"].sum()
    gp = mis_report[(mis_report["Category"] == "Group Participants")
                    & (mis_report["Program"].isin(["FAME", "FAME Group", "CBT Adult Group", "CBT Senior Group"]))] \
                    .groupby("Program", as_index=False)["Total"].sum()

    cbt_sessions = _get_program_value(gs, "CBT Adult Group")
    fame_sessions = _get_program_value(gs, "FAME")
    famegrp_sessions = _get_program_value(gs, "FAME Group")

    cbt_participants = _get_program_value(gp, "CBT Adult Group")
    fame_participants = _get_program_value(gp, "FAME")
    famegrp_participants = _get_program_value(gp, "FAME Group")

    # Caregiver Support from visits
    cgs = visits[visits["ServiceCode"].isin(["Caregiver Support"])].copy()
    cg_participants = float(nunique_safe(cgs["AlaycareID"]))
    cg_visits = float(cgs["Visit Visit ID"].nunique())

    return CaregiverSupportTables(
        cbt_sessions, fame_sessions, famegrp_sessions,
        cbt_participants, fame_participants, famegrp_participants,
        cg_visits, cg_participants
    )


# =========================
#       CSV "WRITER"
# =========================
class SheetWriter:
    """
    CSV-backed writer with the same API used by write_report().
    Keeps the layout identical to your original Excel output:
    - First row: Title in col 0, headers in col 1..N
    - Data rows: Label in col 0, numbers in col 1..N
    - Blank row between sections
    """
    def __init__(self, out_path: Path):
        self.out_path = out_path
        self.table: list[list[Optional[str]]] = []
        self.row = 0

    def _ensure_row(self, r: int, cols: int):
        while len(self.table) <= r:
            self.table.append([])
        # pad row to desired number of columns
        cur = len(self.table[r])
        if cur < cols:
            self.table[r].extend([None] * (cols - cur))

    def _normalize(self):
        # Normalize column widths and convert None -> ""
        max_cols = max((len(r) for r in self.table), default=0)
        for r in self.table:
            if len(r) < max_cols:
                r.extend([None] * (max_cols - len(r)))
        for r in range(len(self.table)):
            self.table[r] = [("" if v is None else v) for v in self.table[r]]

    def finalize(self) -> pd.DataFrame:
        """Write CSV and return the DataFrame (for Output table)."""
        self._normalize()
        df = pd.DataFrame(self.table)
        df.to_csv(self.out_path, index=False, header=False, encoding="utf-8-sig")
        return df

    def write_section(self, title: str, headers: Sequence[str], row_labels: Sequence[str]) -> int:
        # Title + headers row
        cols = 1 + len(headers)
        self._ensure_row(self.row, cols)
        self.table[self.row][0] = title
        for j, h in enumerate(headers, start=1):
            self.table[self.row][j] = h
        self.row += 1

        start = self.row
        # Data rows with labels and blanks
        for lab in row_labels:
            self._ensure_row(self.row, cols)
            self.table[self.row][0] = lab
            # numeric cells left as None
            self.row += 1

        # Blank spacer row
        self._ensure_row(self.row, 1)
        self.row += 1

        return start

    def write_num(self, r: int, c: int, val: Optional[float]):
        if val is None:
            return
        try:
            v = float(val)
            if math.isnan(v):
                return
            v_int = int(round(v))          # integers for non-Remaining rows
            self._ensure_row(r, c + 1)
            self.table[r][c] = v_int
        except Exception:
            pass

    def write_dec(self, r: int, c: int, val: Optional[float]):
        """Write a percentage like 40.00% when val is a fraction (e.g., 0.4)."""
        if val is None:
            return
        try:
            v = float(val)
            self._ensure_row(r, c + 1)
            self.table[r][c] = f"{v*100:.2f}%"
        except Exception:
            pass



def write_report(
    out_path: Path,
    targets: Dict[str, Dict[str, Optional[float]]],
    meals: MealsResult,
    day: DayServicesResult,                      # <-- NEW
    soc: SocialCongregateResult,
    comb: CombinedPSHMRespiteResult,
    als: AssistedLivingResult,
    cm: CaseManagementResult,
    cg: CaregiverSupportTables
) -> pd.DataFrame:
    def tgt(program: str, measure: str) -> Optional[float]:
        return targets.get(program, {}).get(_norm_measure(measure))

    sw = SheetWriter(out_path)

    # ---- Section 1: Meals Delivery ----
    md_headers = ["Individuals Served", "Meal Delivered Combined"]
    md_rows = ["Adult Day Program", "Meals on Wheel", "Year to Date Total",
               "Meals Delivery Fiscal Year Target", "Target Remaining"]
    md = sw.write_section("CSS IH Meals Delivery", md_headers, md_rows)

    # Program rows
    sw.write_num(md + 0, 1, meals.adp_individuals)
    sw.write_num(md + 0, 2, meals.adp_meals)
    sw.write_num(md + 1, 1, meals.mow_individuals)
    sw.write_num(md + 1, 2, meals.mow_meals_weighted)

    # YTD
    md_ytd_ind  = sum_none(meals.adp_individuals, meals.mow_individuals)
    md_ytd_meal = sum_none(meals.adp_meals, meals.mow_meals_weighted)
    sw.write_num(md + 2, 1, md_ytd_ind)
    sw.write_num(md + 2, 2, md_ytd_meal)

    # Targets & Remaining (as percentages)
    md_tgt_ind  = tgt("CSS IH Meals Delivery", "Individuals Served")
    md_tgt_meal = tgt("CSS IH Meals Delivery", "Meal Delivered Combined")
    sw.write_num(md + 3, 1, md_tgt_ind)
    sw.write_num(md + 3, 2, md_tgt_meal)
    sw.write_dec(md + 4, 1, percent_remaining(md_ytd_ind,  md_tgt_ind))
    sw.write_dec(md + 4, 2, percent_remaining(md_ytd_meal, md_tgt_meal))

    # ---- Section 2: Day Services (NEW top-level) ----
    ds_headers = ["Individuals Served", "Attendance Days"]
    ds_rows = ["Adult Day Program", "Year to Date Total",
               "CSS IH Day Services Fiscal Year Target", "Target Remaining"]
    ds = sw.write_section("CSS IH Day Services", ds_headers, ds_rows)

    # Program row
    sw.write_num(ds + 0, 1, day.adp_individuals)
    sw.write_num(ds + 0, 2, day.adp_attendance)

    # YTD (single program -> same as program values)
    ds_ytd_ind = sum_none(day.adp_individuals)
    ds_ytd_att = sum_none(day.adp_attendance)
    sw.write_num(ds + 1, 1, ds_ytd_ind)
    sw.write_num(ds + 1, 2, ds_ytd_att)

    # Targets & Remaining
    ds_tgt_ind = tgt("CSS IH Day Services", "Individuals Served")
    ds_tgt_att = tgt("CSS IH Day Services", "Attendance Days")
    sw.write_num(ds + 2, 1, ds_tgt_ind)
    sw.write_num(ds + 2, 2, ds_tgt_att)
    sw.write_dec(ds + 3, 1, percent_remaining(ds_ytd_ind, ds_tgt_ind))
    sw.write_dec(ds + 3, 2, percent_remaining(ds_ytd_att, ds_tgt_att))

    # ---- Section 3: Social & Congregate Dining ----
    scd_headers = ["Individuals Served", "Attendance Days"]
    scd_rows = ["Health and Wellness", "York West", "Year to Date Total",
                "Social and Congregate Dining Fiscal Year Target", "Target Remaining"]
    scd = sw.write_section("CSS IH Social and Congregate Dining", scd_headers, scd_rows)

    sw.write_num(scd + 0, 1, soc.health_wellness_individuals)
    sw.write_num(scd + 0, 2, soc.health_wellness_attendance)
    sw.write_num(scd + 1, 1, soc.york_west_individuals)
    sw.write_num(scd + 1, 2, soc.york_west_attendance)

    scd_ytd_ind = sum_none(soc.health_wellness_individuals, soc.york_west_individuals)
    scd_ytd_att = sum_none(soc.health_wellness_attendance, soc.york_west_attendance)
    sw.write_num(scd + 2, 1, scd_ytd_ind)
    sw.write_num(scd + 2, 2, scd_ytd_att)

    scd_tgt_ind = tgt("CSS IH Social and Congregate Dining", "Individuals Served")
    scd_tgt_att = tgt("CSS IH Social and Congregate Dining", "Attendance Days")
    sw.write_num(scd + 3, 1, scd_tgt_ind)
    sw.write_num(scd + 3, 2, scd_tgt_att)
    sw.write_dec(scd + 4, 1, percent_remaining(scd_ytd_ind, scd_tgt_ind))
    sw.write_dec(scd + 4, 2, percent_remaining(scd_ytd_att, scd_tgt_att))

    # ---- Section 4: Comb. PS/HM/Respite ----
    comb_headers = ["Individuals Served", "Hours of Care"]
    comb_rows = ["Home Help", "Personal Support Worker(RCVIS, RPOC)", "Year to Date Total",
                 "Comb. PS/HM/Respite Fiscal Year Target", "Target Remaining"]
    cs = sw.write_section("CSS IH Comb. PS/HM/Respite", comb_headers, comb_rows)

    sw.write_num(cs + 0, 1, comb.home_help_individuals)
    sw.write_num(cs + 0, 2, comb.home_help_hours)
    sw.write_num(cs + 1, 1, comb.psw_respite_individuals)
    sw.write_num(cs + 1, 2, comb.psw_respite_hours)

    comb_ytd_ind = sum_none(comb.home_help_individuals, comb.psw_respite_individuals)
    comb_ytd_hrs = sum_none(comb.home_help_hours, comb.psw_respite_hours)
    sw.write_num(cs + 2, 1, comb_ytd_ind)
    sw.write_num(cs + 2, 2, comb_ytd_hrs)

    comb_tgt_ind = tgt("CSS IH Comb. PS/HM/Respite", "Individuals Served")
    comb_tgt_hrs = tgt("CSS IH Comb. PS/HM/Respite", "Hours of Care")
    sw.write_num(cs + 3, 1, comb_tgt_ind)
    sw.write_num(cs + 3, 2, comb_tgt_hrs)
    sw.write_dec(cs + 4, 1, percent_remaining(comb_ytd_ind, comb_tgt_ind))
    sw.write_dec(cs + 4, 2, percent_remaining(comb_ytd_hrs, comb_tgt_hrs))

    # ---- Section 5: Assisted Living Services ----
    als_headers = ["Individuals Served", "Hours of Care", "Inpatient/Resident Days"]
    als_rows = ["Personal Support Worker", "Year to Date Total",
                "CSS IH Assisted Living Services Fiscal Year Target", "Target Remaining"]
    alsr = sw.write_section("CSS IH Assisted Living Services", als_headers, als_rows)

    sw.write_num(alsr + 0, 1, als.psw_individuals)
    sw.write_num(alsr + 0, 2, als.psw_hours)
    sw.write_num(alsr + 0, 3, als.resident_days)

    # Single program YTD = same values
    sw.write_num(alsr + 1, 1, als.psw_individuals)
    sw.write_num(alsr + 1, 2, als.psw_hours)
    sw.write_num(alsr + 1, 3, als.resident_days)

    als_tgt_ind = tgt("CSS IH Assisted Living Services", "Individuals Served")
    als_tgt_hrs = tgt("CSS IH Assisted Living Services", "Hours of Care")
    als_tgt_res = tgt("CSS IH Assisted Living Services", "Inpatient/Resident Days")
    sw.write_num(alsr + 2, 1, als_tgt_ind)
    sw.write_num(alsr + 2, 2, als_tgt_hrs)
    sw.write_num(alsr + 2, 3, als_tgt_res)
    sw.write_dec(alsr + 3, 1, percent_remaining(als.psw_individuals, als_tgt_ind))
    sw.write_dec(alsr + 3, 2, percent_remaining(als.psw_hours,     als_tgt_hrs))
    sw.write_dec(alsr + 3, 3, percent_remaining(als.resident_days, als_tgt_res))

    # ---- Section 6: Case Management ----
    cm_headers = ["Individuals Served", "Visits"]
    cm_rows = ["Case Management", "Year to Date Total",
               "CSS Case Management Fiscal Year Target", "Target Remaining"]
    cmr = sw.write_section("CSS Case Management", cm_headers, cm_rows)

    sw.write_num(cmr + 0, 1, cm.individuals)
    sw.write_num(cmr + 0, 2, cm.visits_count)
    sw.write_num(cmr + 1, 1, cm.individuals)
    sw.write_num(cmr + 1, 2, cm.visits_count)

    cm_tgt_ind = tgt("CSS Case Management", "Individuals Served")
    cm_tgt_vis = tgt("CSS Case Management", "Visits")
    sw.write_num(cmr + 2, 1, cm_tgt_ind)
    sw.write_num(cmr + 2, 2, cm_tgt_vis)
    sw.write_dec(cmr + 3, 1, percent_remaining(cm.individuals, cm_tgt_ind))
    sw.write_dec(cmr + 3, 2, percent_remaining(cm.visits_count, cm_tgt_vis))

    # ---- Section 7: Caregiver Support ----
    cg_headers = ["Individuals Served (Group Participant)", "Visits", "Group Sessions", "SPGIs"]
    cg_rows = ["CBT Adult Group", "FAME", "FAME Group", "Caregiver Support",
               "Year to Date Total", "CSS IH Caregiver Support Fiscal Year Target", "Target Remaining"]
    cgr = sw.write_section("CSS IH Caregiver Support", cg_headers, cg_rows)

    # Program rows
    sw.write_num(cgr + 0, 1, cg.cbt_participants)
    sw.write_num(cgr + 0, 3, cg.cbt_sessions)
    sw.write_num(cgr + 0, 4, cg.cbt_sessions)  # SPGIs = Sessions

    sw.write_num(cgr + 1, 1, cg.fame_participants)
    sw.write_num(cgr + 1, 3, cg.fame_sessions)
    sw.write_num(cgr + 1, 4, cg.fame_sessions)  # SPGIs = Sessions

    sw.write_num(cgr + 2, 1, cg.famegrp_participants)
    sw.write_num(cgr + 2, 3, cg.famegrp_sessions)
    sw.write_num(cgr + 2, 4, cg.famegrp_sessions)  # SPGIs = Sessions

    # Caregiver Support (overall) row
    sw.write_num(cgr + 3, 1, cg.cg_participants)
    sw.write_num(cgr + 3, 2, cg.cg_visits)

    # YTD (include CBT/FAME/FAME Group + Caregiver Support overall row)
    cg_ytd_part = sum_none(cg.cbt_participants, cg.fame_participants, cg.famegrp_participants, cg.cg_participants)
    cg_ytd_vis  = sum_none(cg.cg_visits)
    cg_ytd_sess = sum_none(cg.cbt_sessions, cg.fame_sessions, cg.famegrp_sessions)
    sw.write_num(cgr + 4, 1, cg_ytd_part)
    sw.write_num(cgr + 4, 2, cg_ytd_vis)
    sw.write_num(cgr + 4, 3, cg_ytd_sess)
    sw.write_num(cgr + 4, 4, cg_ytd_sess)  # SPGIs YTD = Sessions YTD

    # Targets & Remaining (as percentages)
    cg_tgt_part = tgt("CSS IH Caregiver Support", "Individuals Served")
    cg_tgt_vis  = tgt("CSS IH Caregiver Support", "Visits")
    cg_tgt_sess = tgt("CSS IH Caregiver Support", "Group Sessions")
    cg_tgt_spgi = tgt("CSS IH Caregiver Support", "SPGIs")
    sw.write_num(cgr + 5, 1, cg_tgt_part)
    sw.write_num(cgr + 5, 2, cg_tgt_vis)
    sw.write_num(cgr + 5, 3, cg_tgt_sess)
    sw.write_num(cgr + 5, 4, cg_tgt_spgi)

    sw.write_dec(cgr + 6, 1, percent_remaining(cg_ytd_part, cg_tgt_part))
    sw.write_dec(cgr + 6, 2, percent_remaining(cg_ytd_vis,  cg_tgt_vis))
    sw.write_dec(cgr + 6, 3, percent_remaining(cg_ytd_sess, cg_tgt_sess))
    sw.write_dec(cgr + 6, 4, percent_remaining(cg_ytd_sess, cg_tgt_spgi))  # SPGIs uses the same actuals as sessions

    # Write CSV and return DF for the Output table
    return sw.finalize()


# =========================
#            RUN
# =========================
def run(output_path: Optional[Path | str] = None):
    """
    Execute the pipeline, write the CSV, and return the DataFrame
    (so the app can show it in the Output table).
    """
    out_path = Path(output_path) if output_path is not None else OUTPUT_XLSX

    # 1) Load source data
    visits = read_csv(ALAYACARE_CLEAN / "Visits.csv")
    adt_df = read_csv(ALAYACARE_CLEAN / "ADT with Zone and Reference.csv")
    mis_report = read_csv(TREAT_CLEAN / "rpt_MIS_Stats.csv")
    targets = load_targets(TARGETS_CSV)

    # 2) Date filtering for visits
    visits["Date"] = coerce_datetime(visits["Date"])
    start_date, end_date = fiscal_window_until_last_month()
    visits = visits.loc[visits["Date"].between(start_date, end_date)].copy()

    # Ensure numeric Duration for hour sums
    if "Duration" in visits.columns:
        visits["Duration"] = pd.to_numeric(visits["Duration"], errors="coerce").fillna(0)

    # 3) Compute sections
    today = pd.Timestamp.today().normalize()
    fy_start = pd.Timestamp(
        year=today.year if today.month >= FY_START_MONTH else today.year - 1,
        month=FY_START_MONTH, day=1
    )

    meals = compute_meals(visits)
    day   = compute_day_services(visits)      # NEW section
    soc   = compute_social_congregate(visits) # (no ADS inside now)
    comb  = compute_combined_ps_hm_respite(visits)
    als   = compute_assisted_living(adt_df, visits, today, fy_start)
    cm    = compute_case_management(visits)
    cg    = compute_caregiver_support_tables(mis_report, visits)

    # (Optional) console sanity prints
    print("[Meals] ADP indiv:", meals.adp_individuals, "| ADP meals:", meals.adp_meals,
          "| MOW indiv:", meals.mow_individuals, "| MOW meals (weighted):", meals.mow_meals_weighted)
    print("[Day Services] ADP indiv:", day.adp_individuals, "| ADP attendance:", day.adp_attendance)
    print("[Social] H&W indiv:", soc.health_wellness_individuals, "| H&W att:", soc.health_wellness_attendance,
          "| YorkW indiv:", soc.york_west_individuals, "| YorkW att:", soc.york_west_attendance)
    print("[Comb] HH indiv:", comb.home_help_individuals, "| HH hours:", comb.home_help_hours,
          "| PSW Respite indiv:", comb.psw_respite_individuals, "| PSW Respite hours:", comb.psw_respite_hours)
    print("[ALS] PSW indiv:", als.psw_individuals, "| PSW hours:", als.psw_hours, "| Resident days:", als.resident_days)
    print("[CM] indiv:", cm.individuals, "| visits:", cm.visits_count)
    print("[CG] CBT sessions:", cg.cbt_sessions, "| FAME sessions:", cg.fame_sessions, "| FAMEGrp sessions:", cg.famegrp_sessions,
          "| CBT part:", cg.cbt_participants, "| FAME part:", cg.fame_participants, "| FAMEGrp part:", cg.famegrp_participants,
          "| CG visits:", cg.cg_visits, "| CG participants:", cg.cg_participants)

    # 4) Build report (writes CSV and returns DF for Output table)
    df_out = write_report(
        out_path,
        targets,
        meals,
        day,    # NEW
        soc,
        comb,
        als,
        cm,
        cg
    )
    print(f"Saved: {out_path}")

    # Replace NaN/None with blanks for nice display in Output table
    return df_out.fillna("")


# =========================
#           MAIN
# =========================
def main() -> None:
    _ = run()

if __name__ == "__main__":
    main()
