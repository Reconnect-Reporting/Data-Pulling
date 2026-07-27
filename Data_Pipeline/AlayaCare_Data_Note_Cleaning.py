# -*- coding: utf-8 -*-
import re
from pathlib import Path
from datetime import datetime, timedelta
import shutil
import glob

# ===== Paths (auto-detect username) =====
OD = Path.home() / "OneDrive - Reconnect Community Health Services (1)"
RAW_BASE = OD / "data" / "Raw Data" / "AlayaCare"
CLEAN_DIR = OD / "data" / "Clean Data" / "AlayaCare"
CLEAN_DIR.mkdir(parents=True, exist_ok=True)

# Raw folder -> base name for output (rename to *_daily)
TARGETS = {
    "notes_client_look":   "notes_client_look_daily",
    "notes_progress_look": "notes_progress_look_daily",
    "notes_visit_look":    "notes_visit_look_daily",
}

EXTS = (".csv", ".xlsx", ".xls")
DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")  # e.g., 2025-11-13


def extract_date_from_name(name: str) -> datetime | None:
    m = DATE_RE.search(name)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y-%m-%d")
    except ValueError:
        return None


def find_most_recent_file_by_date(folder: Path, stem_prefix: str) -> Path | None:
    """
    If today is Monday → include last 3 days (Fri, Sat, Sun)
    Otherwise → just take latest file as usual
    """
    today = datetime.today()
    is_monday = today.weekday() == 0  # Monday = 0

    candidates = []

    for ext in EXTS:
        pattern = str(folder / f"{stem_prefix}_*{ext}")
        for f in glob.glob(pattern):
            p = Path(f)
            dt = extract_date_from_name(p.name)
            if dt:
                # NEW LOGIC
                if is_monday:
                    if (today - dt).days <= 3:
                        candidates.append((p, dt, p.stat().st_mtime))
                else:
                    candidates.append((p, dt, p.stat().st_mtime))

    if not candidates:
        return None

    # Sort by (date, mtime) descending
    candidates.sort(key=lambda x: (x[1], x[2]), reverse=True)
    return candidates[0][0]


def atomic_move_replace(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".tmp")

    if tmp.exists():
        tmp.unlink()

    shutil.copy2(src, tmp)

    if dst.exists():
        dst.unlink()

    tmp.rename(dst)

    try:
        src.unlink()
    except Exception:
        pass


def process_one(raw_folder_name: str, output_base: str):
    raw_folder = RAW_BASE / raw_folder_name
    if not raw_folder.exists():
        print(f"[{raw_folder_name}] Raw folder not found: {raw_folder}")
        return

    latest = find_most_recent_file_by_date(raw_folder, raw_folder_name)
    if not latest:
        print(f"[{raw_folder_name}] No valid files found")
        return

    dest = CLEAN_DIR / f"{output_base}{latest.suffix}"

    try:
        atomic_move_replace(latest, dest)
        print(f"[{raw_folder_name}] Moved {latest.name} -> {dest}")
    except Exception as e:
        print(f"[{raw_folder_name}] ERROR moving {latest} -> {dest}: {e}")


def main():
    for raw_folder_name, output_base in TARGETS.items():
        process_one(raw_folder_name, output_base)


if __name__ == "__main__":
    main()