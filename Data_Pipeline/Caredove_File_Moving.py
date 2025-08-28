import os
import shutil
from pathlib import Path

# ---------- Helpers ----------
def get_downloads_folder() -> Path:
    """
    Get the current user's Downloads folder using Windows Known Folders.
    Falls back to ~/Downloads if the API call fails.
    """
    try:
        import ctypes
        from ctypes import wintypes

        # FOLDERID_Downloads GUID
        FOLDERID_Downloads = ctypes.c_char_p(b"{374DE290-123F-4565-9164-39C4925E467B}")

        _SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
        _SHGetKnownFolderPath.argtypes = [ctypes.c_void_p, ctypes.wintypes.DWORD, ctypes.wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)]  # type: ignore

        path_ptr = ctypes.c_wchar_p()
        # Flags=0, hToken=None
        res = _SHGetKnownFolderPath(ctypes.c_void_p(int(FOLDERID_Downloads.value.decode("ascii").strip("{}"), 16)), 0, None, ctypes.byref(path_ptr))
        if res == 0 and path_ptr.value:
            return Path(path_ptr.value)
    except Exception:
        pass

    # Fallback
    return Path.home() / "Downloads"


def get_onedrive_root() -> Path:
    """
    Try the standard OneDrive environment variables.
    If not found, fall back to the first directory in the user's home that starts with 'OneDrive'.
    """
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.environ.get(var)
        if p and Path(p).exists():
            return Path(p)

    # Fallback: find any OneDrive-like folder in home
    home = Path.home()
    candidates = [d for d in home.iterdir() if d.is_dir() and d.name.startswith("OneDrive")]
    if candidates:
        return candidates[0]

    raise FileNotFoundError("Could not find a OneDrive folder. Ensure OneDrive is installed/signed in.")


def move_all(src_dir: Path, dst_dir: Path) -> int:
    """
    Move all non-temp files from src_dir to dst_dir.
    Overwrites existing files at the destination.
    Returns the count of files moved.
    """
    if not src_dir.exists():
        print(f"Source folder does not exist: {src_dir}")
        return 0

    dst_dir.mkdir(parents=True, exist_ok=True)

    skip_exts = {".crdownload", ".part", ".tmp"}
    moved = 0

    for p in src_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() in skip_exts:
            continue

        target = dst_dir / p.name
        if target.exists():
            target.unlink()  # overwrite
        shutil.move(str(p), str(target))
        print(f"Moved: {p.name}  →  {target}")
        moved += 1

    return moved


def run():
    """Entry point for the pipeline (imported and called by main.py)."""
    downloads = get_downloads_folder()
    caredove_src = downloads / "caredove"

    onedrive_root = get_onedrive_root()
    caredove_dst = onedrive_root / "data" / "Raw Data" / "Caredove"

    print(f"Source:      {caredove_src}")
    print(f"Destination: {caredove_dst}")

    count = move_all(caredove_src, caredove_dst)
    print(f"Done. Files moved: {count}")
    return count


# ---------- Main ----------
if __name__ == "__main__":
    downloads = get_downloads_folder()
    caredove_src = downloads / "caredove"

    onedrive_root = get_onedrive_root()
    caredove_dst = onedrive_root / "data" / "Raw Data" / "Caredove"

    print(f"Source:      {caredove_src}")
    print(f"Destination: {caredove_dst}")

    count = move_all(caredove_src, caredove_dst)
    print(f"Done. Files moved: {count}")
