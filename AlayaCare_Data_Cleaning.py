# -*- coding: utf-8 -*-
import os
import glob
import pandas as pd
from pathlib import Path

# ===== Paths =====
def _onedrive_root() -> Path:
    # Prefer Windows env vars (most reliable for OneDrive for Business)
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.environ.get(var)
        if p and Path(p).exists():
            return Path(p)
    # Fallback: first "OneDrive*" folder under the user profile
    home = Path.home()
    for child in home.iterdir():
        if child.is_dir() and child.name.startswith("OneDrive"):
            return child
    # Last resort
    return home / "OneDrive"

OD = _onedrive_root()
DATA = OD / "data"

CLEAN_DIR            = str(DATA / "Clean Data" / "AlayaCare")
RAW_VISITS_DIR       = str(DATA / "Raw Data" / "AlayaCare" / "Visits")
RAW_NOTES_DIR        = str(DATA / "Raw Data" / "AlayaCare" / "Note")
RAW_CLIENTCALLS_DIR  = str(DATA / "Raw Data" / "AlayaCare" / "Client Calls")

# ===== Helpers =====
def is_temp_or_hidden(path: str) -> bool:
    base = os.path.basename(path)
    return base.startswith("~$") or base.startswith("._")

def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
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

def write_atomic_csv(df: pd.DataFrame, final_path: str):
    tmp_path = final_path + ".tmp"
    df.to_csv(tmp_path, index=False, encoding="utf-8-sig")
    if os.path.exists(final_path):
        os.remove(final_path)
    os.replace(tmp_path, final_path)

def write_atomic_xlsx(df: pd.DataFrame, final_path: str, sheet_name: str):
    tmp_path = final_path + ".tmp"
    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    if os.path.exists(final_path):
        os.remove(final_path)
    os.replace(tmp_path, final_path)

def find_existing_main(base_names: list[str]) -> str | None:
    """
    Return path of existing main file if found, preferring CSV, then XLSX, then XLS.
    base_names: list of candidate base names to try (e.g., ["Notes", "Note"])
    """
    exts = [".csv", ".xlsx", ".xls"]
    for base in base_names:
        for ext in exts:
            p = os.path.join(CLEAN_DIR, base + ext)
            if os.path.exists(p):
                return p
    return None

def gather_raw_files(raw_dir: str, patterns: list[str]) -> list[str]:
    files = []
    for pat in patterns:
        files.extend(glob.glob(os.path.join(raw_dir, pat)))
    return [p for p in files if not is_temp_or_hidden(p)]

def process_dataset(
    label: str,
    raw_dir: str,
    raw_patterns: list[str],
    main_base_candidates: list[str],
    default_main_basename: str,
    sheet_name: str
):
    """
    Concatenate raw files + existing main (if present),
    drop full-row duplicates, and replace the main table.
    """
    os.makedirs(CLEAN_DIR, exist_ok=True)

    daily_files = gather_raw_files(raw_dir, raw_patterns)
    main_path = find_existing_main(main_base_candidates)

    if not daily_files and not main_path:
        print(f"[{label}] No raw files and no main file found. Skipping.")
        return

    frames = []

    # Include existing main first (so new raw rows still kept after drop_duplicates)
    if main_path:
        try:
            df_main = read_table(main_path)
            frames.append(df_main)
            print(f"[{label}] Loaded main: {os.path.basename(main_path)} -> {len(df_main):,} rows")
        except Exception as e:
            print(f"[{label}] WARNING: Failed to read main ({main_path}): {e}")

    # Read raw files
    for p in sorted(daily_files):
        try:
            df = read_table(p)
            frames.append(df)
            print(f"[{label}] Loaded raw: {os.path.basename(p)} -> {len(df):,} rows")
        except Exception as e:
            print(f"[{label}] WARNING: Skipped {p} due to read error: {e}")

    if not frames:
        print(f"[{label}] No readable data. Skipping.")
        return

    combined = pd.concat(frames, ignore_index=True, sort=False)
    before = len(combined)
    combined = combined.drop_duplicates()
    after = len(combined)
    print(f"[{label}] Concatenated: {before:,} -> after drop_duplicates: {after:,}")

    # Decide output path/format: keep existing format if present; else default to CSV with default_main_basename
    if main_path:
        base_no_ext, ext = os.path.splitext(main_path)
        out_path = main_path  # overwrite in place
        write_as_xlsx = ext.lower() in (".xlsx", ".xls")
    else:
        out_path = os.path.join(CLEAN_DIR, default_main_basename + ".csv")
        write_as_xlsx = False

    # Write atomically
    if write_as_xlsx:
        write_atomic_xlsx(combined, out_path, sheet_name=sheet_name)
    else:
        write_atomic_csv(combined, out_path)

    print(f"[{label}] Updated main written to: {out_path}")

# ===== Run all three =====
def main():
    # Visits
    process_dataset(
        label="Visits",
        raw_dir=RAW_VISITS_DIR,
        raw_patterns=["Visit*.csv", "Visit*.xlsx", "Visit*.xls"],
        main_base_candidates=["Visits", "Visit"],     # prefer 'Visits', but handle 'Visit' if that's how it was saved
        default_main_basename="Visits",
        sheet_name="Visits"
    )

    # Notes
    process_dataset(
        label="Notes",
        raw_dir=RAW_NOTES_DIR,
        raw_patterns=["Note*.csv", "Note*.xlsx", "Note*.xls", "Notes*.csv", "Notes*.xlsx", "Notes*.xls"],
        main_base_candidates=["Notes", "Note"],       # prefer 'Notes', handle 'Note' if present
        default_main_basename="Notes",
        sheet_name="Notes"
    )

    # Client Calls
    process_dataset(
        label="Client Calls",
        raw_dir=RAW_CLIENTCALLS_DIR,
        raw_patterns=[
            "Client Calls*.csv", "Client Calls*.xlsx", "Client Calls*.xls",
            "ClientCalls*.csv", "ClientCalls*.xlsx", "ClientCalls*.xls",
            "Client_Calls*.csv", "Client_Calls*.xlsx", "Client_Calls*.xls",
        ],
        main_base_candidates=["Client Calls", "ClientCalls", "Client_Calls"],
        default_main_basename="Client Calls",
        sheet_name="Client Calls"
    )

if __name__ == "__main__":
    main()
