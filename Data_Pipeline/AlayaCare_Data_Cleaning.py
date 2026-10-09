# -*- coding: utf-8 -*-
import os, re, glob, shutil
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
from calendar import month_name


# ===== Paths =====
OD = Path.home() / "OneDrive - Reconnect Community Health Services (1)"
RAW_ROOT   = OD / "data" / "Raw Data" / "AlayaCare"
CLEAN_DIR  = OD / "data" / "Clean Data" / "AlayaCare"
CLEAN_DIR.mkdir(parents=True, exist_ok=True)

CLIENT_LOOKUP_DIR = (
    Path.home()
    / "Reconnect Community Health Services"
    / "Health Services - Documents"
    / "_library"
    / "client_lookup"
)

# ===== Config =====
EXTS    = [".csv", ".xlsx", ".xls"]
DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")
APPEND_DAYS = 5  # last 5 days

# Raw folder -> Clean fixed base filename (no date)
LATEST_COPY_REPORTS = {
    "ADT_with_Zone_and_Reference": "ADT with Zone and Reference",
    "ADT_with_Zone_and_Reference_New_Referral_Date": "ADT with Zone and Reference New Referral Date",
    "CM_Supervisors":              "CM Supervisors",
    "CM_Supervisors_Discharged":   "CM Supervisors_Discharged",
    "Heart_at_Home":               "Heart at Home",
    "HeartatHomeAddresss":         "HeartatHomeAddress",
    "HHRI_VHA":                    "HHRI(VHA)",
    "Meals_on_Wheels":             "Meals_on_Wheels",
}

# Append mode (collect last N days and merge into main table)
# Keys = raw folder names in RAW_ROOT; values = clean base filenames (no date).
APPEND_REPORTS = {
    "Client_Calls":          "Client Calls",          # generic append
    "Facilities_Attendance": "Facilities Attendance", # generic append
    "Form_Report":           "Form Report",           # generic append
    "Notes":                 "Notes",                 # custom: append + drop_duplicates
    "Visits":                "Visits",   
    "notes_client_look":       "notes_client_look",
    "notes_progress_look":      "notes_progress_look",
    "notes_visit_look": "notes_visit_look"              # custom: keep newest on key (see below)
}

# ===== Helpers =====

def list_all_files(folder: Path) -> list[Path]:
    """Return ALL files in folder matching EXTS, sorted by name."""
    hits: list[Path] = []
    for ext in EXTS:
        hits.extend(Path(folder).glob(f"*{ext}"))
    return sorted(hits)

def previous_month_name_from_date(dt: datetime) -> str:
    """Return the full English month name for the month BEFORE dt."""
    first_of_month = dt.replace(day=1)
    prev_month_day = first_of_month - timedelta(days=1)
    return prev_month_day.strftime("%B")

def extract_date_from_name(name: str) -> datetime | None:
    m = DATE_RE.search(name)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y-%m-%d")
    except ValueError:
        return None

def find_latest_file(folder: Path) -> Path | None:
    candidates = []
    for ext in EXTS:
        for f in glob.glob(str(folder / f"*{ext}")):
            dt = extract_date_from_name(os.path.basename(f))
            mtime = os.path.getmtime(f)
            candidates.append((Path(f), dt, mtime))
    if not candidates:
        return None
    # Sort: latest explicit date first; if no date, use mtime
    candidates.sort(key=lambda x: ((x[1] is not None, x[1] or datetime.min), x[2]), reverse=True)
    return candidates[0][0]

def atomic_copy(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".tmp")
    shutil.copy2(src, tmp)
    if dst.exists():
        dst.unlink()
    tmp.rename(dst)

def read_table(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext in (".csv", ".txt"):
        for enc in ("utf-8-sig", "utf-8", "cp1252"):
            try:
                return pd.read_csv(path, encoding=enc)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path, errors="replace")
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path, sheet_name=0)
    else:
        raise ValueError(f"Unsupported file type: {path}")

def write_atomic(df: pd.DataFrame, dst: Path, sheet_name: str = "Sheet1"):
    tmp = dst.with_suffix(dst.suffix + ".tmp")
    if dst.suffix.lower() in (".xlsx", ".xls"):
        with pd.ExcelWriter(tmp, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name=sheet_name)
    else:
        df.to_csv(tmp, index=False, encoding="utf-8-sig")
    if dst.exists():
        dst.unlink()
    tmp.rename(dst)

def list_recent_files(folder: Path, days: int) -> list[Path]:
    """Pick files whose filename date is within last `days` (inclusive). If no date, fallback to mtime window."""
    cutoff = datetime.now() - timedelta(days=days)
    hits: list[Path] = []
    for ext in EXTS:
        for f in glob.glob(str(folder / f"*{ext}")):
            p = Path(f)
            dt = extract_date_from_name(p.name)
            if dt:
                if dt >= cutoff:
                    hits.append(p)
            else:
                if datetime.fromtimestamp(p.stat().st_mtime) >= cutoff:
                    hits.append(p)
    return sorted(hits)

def align_columns_like(left: pd.DataFrame, right: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Union columns, fill missing with NA, order as in left then new ones."""
    left_cols  = list(left.columns)
    right_cols = list(right.columns)
    all_cols = list(dict.fromkeys(left_cols + right_cols))  # keep order
    left2 = left.copy()
    right2 = right.copy()
    for c in all_cols:
        if c not in left2.columns:
            left2[c] = pd.NA
        if c not in right2.columns:
            right2[c] = pd.NA
    return left2[all_cols], right2[all_cols]

# ===== Merge strategies =====
def merge_generic_append(clean_name: str, existing: Path, recent_files: list[Path]) -> pd.DataFrame | None:
    dfs = []
    if existing.exists():
        try:
            base_df = read_table(existing)
            dfs.append(base_df)
            print(f"[{clean_name}] Loaded main: {existing.name} ({len(base_df):,} rows)")
        except Exception as e:
            print(f"[{clean_name}] WARNING: failed to read main ({existing}): {e}")

    for p in recent_files:
        try:
            dfp = read_table(p)
            dfs.append(dfp)
            print(f"[{clean_name}] + {p.name} ({len(dfp):,} rows)")
        except Exception as e:
            print(f"[{clean_name}] WARNING: skip {p.name}: {e}")

    if not dfs:
        return None
    merged = pd.concat(dfs, ignore_index=True, sort=False).drop_duplicates()
    return merged

def merge_notes_append(clean_name: str, existing: Path, recent_files: list[Path]) -> pd.DataFrame | None:
    """Notes: append and drop_duplicates (simple)."""
    return merge_generic_append(clean_name, existing, recent_files)

def merge_visit_keep_new(clean_name: str, existing: Path, recent_files: list[Path]) -> pd.DataFrame | None:
    """
    Visits: drop service/client status columns, then de-dupe on all columns
    EXCEPT 'Visit Status', keeping the newly appended rows when there is a match.
    """
    base = pd.DataFrame()
    if existing.exists():
        try:
            base = read_table(existing)
            print(f"[{clean_name}] Loaded main: {existing.name} ({len(base):,} rows)")
        except Exception as e:
            print(f"[{clean_name}] WARNING: failed to read main ({existing}): {e}")
            base = pd.DataFrame()

    new_parts = []
    for p in recent_files:
        try:
            dfp = read_table(p)
            new_parts.append(dfp)
            print(f"[{clean_name}] + {p.name} ({len(dfp):,} rows)")
        except Exception as e:
            print(f"[{clean_name}] WARNING: skip {p.name}: {e}")

    if not new_parts and base.empty:
        return None

    new_df = pd.concat(new_parts, ignore_index=True, sort=False) if new_parts else pd.DataFrame()

    # Drop these on BOTH sides if present (include common header variants)
    drop_cols_candidates = [
        'Service Status', 'ServiceStatus', 'Service_Status',
        'ClientStatus', 'Client Status', 'Client_Status'
    ]
    for c in drop_cols_candidates:
        if c in base.columns:
            base = base.drop(columns=[c])
        if c in new_df.columns:
            new_df = new_df.drop(columns=[c])

    # Align columns
    if base.empty and not new_df.empty:
        base = pd.DataFrame(columns=new_df.columns)
    base, new_df = align_columns_like(base, new_df)

    # Mark provenance and combine
    base['_is_new'] = False
    new_df['_is_new'] = True
    all_rows = pd.concat([base, new_df], ignore_index=True, sort=False)

    # Build key columns = everything except 'Visit Status' and the marker
    exclude = {'Visit Status', '_is_new'}
    # Also exclude common variants of Visit Status name if they appear
    for vs in ['VisitStatus', 'Visit_Status']:
        if vs in all_rows.columns:
            exclude.add(vs)
    keys = [c for c in all_rows.columns if c not in exclude]

    # Prefer new rows on ties
    all_rows = all_rows.sort_values(keys + ['_is_new'])
    result = all_rows.drop_duplicates(subset=keys, keep='last').drop(columns=['_is_new'])

    return result

# ===== Tasks =====

def update_hub_and_spoke_month_label():
    """
    From RAW_ROOT/Hub_and_Spoke pick the latest dated file like:
      Hub_and_Spoke_YYYY-MM-DD.ext
    Use the date in the filename to compute the PREVIOUS MONTH, then copy to:
      CLEAN_DIR/'Hub and Spoke'/Hub_and_Spoke_<MonthName>.ext
    Example: 2025-10-23 -> 'Hub_and_Spoke_September.csv'
    """
    raw_folder = RAW_ROOT / "Hub_and_Spoke"
    if not raw_folder.exists():
        print("[Hub_and_Spoke] Raw folder not found:", raw_folder)
        return

    latest = find_latest_file(raw_folder)
    if not latest:
        print("[Hub_and_Spoke] No files found.")
        return

    dt = extract_date_from_name(latest.name)
    if dt is None:
        # Fallback: use modified time if no date embedded (shouldn't happen per your pattern)
        dt = datetime.fromtimestamp(latest.stat().st_mtime)

    month_label = previous_month_name_from_date(dt)  # e.g., "September"

    dest_dir = CLEAN_DIR / "Hub and Spoke"   # note the space in folder name
    dest_dir.mkdir(parents=True, exist_ok=True)

    dest = dest_dir / f"Hub_and_Spoke_{month_label}{latest.suffix}"
    try:
        atomic_copy(latest, dest)
        print(f"[Hub_and_Spoke] {latest.name} -> {dest.relative_to(CLEAN_DIR)}")
    except Exception as e:
        print(f"[Hub_and_Spoke] ERROR copying {latest}: {e}")

def update_latest_fixed_names():
    for raw_folder_name, clean_base_name in LATEST_COPY_REPORTS.items():
        raw_folder = RAW_ROOT / raw_folder_name
        if not raw_folder.exists():
            print(f"[{clean_base_name}] Raw folder not found: {raw_folder}")
            continue

        latest = find_latest_file(raw_folder)
        if not latest:
            print(f"[{clean_base_name}] No files found in: {raw_folder}")
            continue

        dest = CLEAN_DIR / f"{clean_base_name}{latest.suffix}"
        try:
            atomic_copy(latest, dest)
            print(f"[{clean_base_name}] Replaced Clean with: {latest.name} -> {dest.name}")

            # ADT with Zone and Reference only:
            # also overwrite a second copy in client_lookup.
            if clean_base_name == "ADT with Zone and Reference":
                CLIENT_LOOKUP_DIR.mkdir(parents=True, exist_ok=True)
                second_dest = CLIENT_LOOKUP_DIR / dest.name
                atomic_copy(dest, second_dest)
                print(
                    f"[{clean_base_name}] Also replaced client_lookup copy -> {second_dest}"
                )

        except Exception as e:
            print(f"[{clean_base_name}] ERROR copying {latest}: {e}")

def append_last_n_days():
    SPECIAL_ALL_FILES = {"notes_client_look", "notes_progress_look"}  # <— add this

    for raw_folder_name, clean_base_name in APPEND_REPORTS.items():
        raw_folder = RAW_ROOT / raw_folder_name
        if not raw_folder.exists():
            print(f"[{clean_base_name}] Raw folder not found: {raw_folder}")
            continue

        # Use ALL files for the two notes_* folders; otherwise last N days
        if raw_folder_name in SPECIAL_ALL_FILES:
            recent_files = list_all_files(raw_folder)
        else:
            recent_files = list_recent_files(raw_folder, APPEND_DAYS)

        if not recent_files:
            scope = "ALL files" if raw_folder_name in SPECIAL_ALL_FILES else f"last {APPEND_DAYS} days"
            print(f"[{clean_base_name}] No files found ({scope}).")
            continue

        # Decide main Clean path (prefer existing ext; else default to CSV)
        existing = None
        for ext in [".csv", ".xlsx", ".xls"]:
            cand = CLEAN_DIR / f"{clean_base_name}{ext}"
            if cand.exists():
                existing = cand
                break
        if existing is None:
            existing = CLEAN_DIR / f"{clean_base_name}.csv"

        # Branch per special rules
        name_lower = clean_base_name.lower()
        if name_lower == "notes":
            merged = merge_notes_append(clean_base_name, existing, recent_files)
        elif name_lower in ("visit", "visits"):  # <— fix: handle plural
            merged = merge_visit_keep_new(clean_base_name, existing, recent_files)
        else:
            merged = merge_generic_append(clean_base_name, existing, recent_files)

        if merged is None:
            print(f"[{clean_base_name}] Nothing to write.")
            continue

        print(f"[{clean_base_name}] Writing merged: {len(merged):,} rows -> {existing.name}")
        write_atomic(merged, existing, sheet_name=clean_base_name)


# ===== Main =====
def main():
    update_hub_and_spoke_month_label()
    update_latest_fixed_names()  # replace dated reports (drop date in Clean name)
    append_last_n_days()         # append last N days (with custom logic for Notes & Visits)

if __name__ == "__main__":
    main()
