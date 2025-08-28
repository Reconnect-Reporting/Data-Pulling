# move_caredove_files.py
import os
from pathlib import Path
import shutil

ORG_NAME = "Reconnect Community Health Services"
DATA_ROOT = Path("data")
RAW_DIR_REL = Path("Raw Data") / "Caredove"
CLEAN_DIR_REL = Path("Clean Data") / "Caredove"

def _onedrive_root() -> Path:
    """
    Detect OneDrive root robustly:
    - Prefer environment variables
    - Otherwise search in the user's home directory for a folder like:
      OneDrive - Reconnect Community Health Services or OneDrive - ... (1)
    """
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.environ.get(var)
        if p and Path(p).exists():
            return Path(p)

    home = Path.home()
    # Match both with and without "(1)"
    candidates = sorted(home.glob(f"OneDrive - {ORG_NAME}*"), key=lambda p: len(str(p)))
    for c in candidates:
        if c.is_dir():
            return c

    # Last resort: plain OneDrive under home
    default = home / "OneDrive"
    if default.exists():
        return default

    raise FileNotFoundError("Could not find OneDrive folder. Set OneDrive env vars or adjust detection logic.")

def move_overwrite(src: Path, dst: Path) -> None:
    """Move src -> dst, overwriting dst if it exists."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()
    shutil.move(str(src), str(dst))

def run():
    od = _onedrive_root()

    raw_dir   = od / DATA_ROOT / RAW_DIR_REL
    clean_dir = od / DATA_ROOT / CLEAN_DIR_REL

    if not raw_dir.exists():
        raise FileNotFoundError(f"Raw directory not found: {raw_dir}")

    clean_dir.mkdir(parents=True, exist_ok=True)

    moved = 0
    for p in raw_dir.iterdir():
        if p.is_file():
            target = clean_dir / p.name
            move_overwrite(p, target)
            print(f"Moved → {target}")
            moved += 1

    if moved == 0:
        print(f"No files found in: {raw_dir}")
    else:
        print(f"Done. {moved} file(s) moved to {clean_dir}")

if __name__ == "__main__":
    run()
