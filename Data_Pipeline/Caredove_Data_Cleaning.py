# move_caredove_files.py
import os
from pathlib import Path
import shutil

# ---- Org / folder pieces ----
ORG_BASENAME = "Reconnect Community Health Services"
PREFERRED_SUFFIX = " (1)"  # we prefer the folder that includes (1)
DATA_ROOT = Path("data")
RAW_DIR_REL = Path("Raw Data") / "Caredove"
CLEAN_DIR_REL = Path("Clean Data") / "Caredove"


def _onedrive_root() -> Path:

    home = Path.home()

    # 1) Exact preferred folder with (1)
    preferred = home / f"OneDrive - {ORG_BASENAME}{PREFERRED_SUFFIX}"
    if preferred.exists() and preferred.is_dir():
        return preferred

    # 2) Any OneDrive folder that starts with the base name
    candidates = sorted(
        home.glob(f"OneDrive - {ORG_BASENAME}*"),
        key=lambda p: (len(str(p)), p.name.lower())
    )
    if candidates:
        return candidates[0]

    # 3) Env vars
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.environ.get(var)
        if p and Path(p).exists():
            return Path(p)

    # 4) Plain default
    default = home / "OneDrive"
    if default.exists():
        return default

    raise FileNotFoundError(
        "Could not find a OneDrive folder. "
        "Create/Sign in to OneDrive or update _onedrive_root()."
    )


def move_overwrite(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()
    shutil.move(str(src), str(dst))


def run():
    od = _onedrive_root()
    print(f"[info] Using OneDrive root: {od}")

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
