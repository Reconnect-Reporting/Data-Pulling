import os
import shutil
from pathlib import Path
import getpass

# ---------- Helpers ----------
def get_downloads_folder() -> Path:
    """Get current user's Downloads folder; fall back to ~/Downloads."""
    try:
        import ctypes
        from ctypes import wintypes

        FOLDERID_Downloads = ctypes.c_char_p(b"{374DE290-123F-4565-9164-39C4925E467B}")
        _SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
        _SHGetKnownFolderPath.argtypes = [
            ctypes.c_void_p, wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)
        ]
        path_ptr = ctypes.c_wchar_p()
        res = _SHGetKnownFolderPath(
            ctypes.c_void_p(int(FOLDERID_Downloads.value.decode("ascii").strip("{}"), 16)),
            0, None, ctypes.byref(path_ptr)
        )
        if res == 0 and path_ptr.value:
            return Path(path_ptr.value)
    except Exception:
        pass
    return Path.home() / "Downloads"

def get_onedrive_root() -> Path:
    """
    OneDrive root for Reconnect Community Health Services (no trailing '\data').
    We will append 'data/Raw Data/Caredove' downstream.
    """
    user = getpass.getuser()
    return Path(fr"C:\Users\{user}\OneDrive - Reconnect Community Health Services (1)")

def move_all(src_dir: Path, dst_dir: Path) -> int:
    """Move all non-temp files from src_dir to dst_dir (overwrite if exists)."""
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
    # Append 'data' here (root has no trailing 'data' now)
    caredove_dst = onedrive_root / "data" / "Raw Data" / "Caredove"

    print(f"Source:      {caredove_src.resolve()}")
    print(f"Destination: {caredove_dst.resolve()}")

    count = move_all(caredove_src, caredove_dst)
    print(f"Done. Files moved: {count}")
    return count

# ---------- Main ----------
if __name__ == "__main__":
    downloads = get_downloads_folder()
    caredove_src = downloads / "caredove"

    onedrive_root = get_onedrive_root()
    caredove_dst = onedrive_root / "data" / "Raw Data" / "Caredove"

    print(f"Source:      {caredove_src.resolve()}")
    print(f"Destination: {caredove_dst.resolve()}")

    count = move_all(caredove_src, caredove_dst)
    print(f"Done. Files moved: {count}")
