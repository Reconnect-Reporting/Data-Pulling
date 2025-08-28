import os
import time
from pathlib import Path

def get_default_download_dir() -> Path:
    """User's Downloads folder (Windows/macOS/Linux)."""
    return Path(os.path.expanduser("~")) / "Downloads"

def _fmt_bytes(n: int) -> str:
    for unit in ("B","KB","MB","GB","TB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} PB"

def purge_downloads(
    download_dir: Path | None = None,
    recursive: bool = True,
    move_to_recycle_bin: bool = True,   # False = PERMANENT delete
    delete_dirs: bool = False,          # only removes empty dirs if True
    older_than_days: float | None = None,  # e.g., 7 to keep last week’s
):
    """
    Deletes files in Downloads.
    - move_to_recycle_bin=True tries send2trash (Recycle Bin) first, then falls back to permanent if unavailable.
    - recursive=True deletes within subfolders too (files only).
    - older_than_days keeps newer files if set.
    - delete_dirs removes empty subfolders after file deletions when True.
    """
    if download_dir is None:
        download_dir = get_default_download_dir()

    if not download_dir.exists() or not download_dir.is_dir():
        raise RuntimeError(f"Downloads not found: {download_dir}")

    cutoff = None
    if older_than_days is not None:
        cutoff = time.time() - older_than_days * 86400

    # Discover targets
    glob_pat = "**/*" if recursive else "*"
    paths = list(download_dir.glob(glob_pat))

    # Prefer recycle bin if asked
    send2trash = None
    if move_to_recycle_bin:
        try:
            from send2trash import send2trash as _s2t
            send2trash = _s2t
        except Exception:
            send2trash = None  # will do permanent delete

    files_deleted = 0
    bytes_deleted = 0
    errors = []

    for p in paths:
        if p.is_file():
            try:
                if cutoff is not None and p.stat().st_mtime > cutoff:
                    continue  # too new, keep it
                size = p.stat().st_size
                if send2trash:
                    send2trash(str(p))
                else:
                    p.unlink(missing_ok=True)
                files_deleted += 1
                bytes_deleted += size
            except Exception as e:
                errors.append((str(p), repr(e)))

    # Optionally try to remove now-empty folders
    if delete_dirs:
        # remove deepest first
        for p in sorted((pp for pp in paths if pp.is_dir()), key=lambda x: len(x.parts), reverse=True):
            try:
                p.rmdir()
            except Exception:
                pass  # non-empty or in-use

    print(f"Deleted {files_deleted} files (~{_fmt_bytes(bytes_deleted)}). Folder: {download_dir}")
    if move_to_recycle_bin and send2trash:
        print("Files moved to Recycle Bin.")
    elif move_to_recycle_bin and not send2trash:
        print("send2trash not installed — files permanently deleted. Install with: pip install send2trash")
    if errors:
        print("\nSome items could not be removed:")
        for path, err in errors[:10]:
            print(" -", path, "->", err)
        if len(errors) > 10:
            print(f"...and {len(errors)-10} more.")

def run(
    recursive: bool = True,
    move_to_recycle_bin: bool = True,
    delete_dirs: bool = False,
    older_than_days: float | None = None,
):
    """Entry point so main.py can call me."""
    return purge_downloads(
        download_dir=get_default_download_dir(),
        recursive=recursive,
        move_to_recycle_bin=move_to_recycle_bin,
        delete_dirs=delete_dirs,
        older_than_days=older_than_days,
    )

# Optional: keep a CLI entry too
def main():
    return run()

if __name__ == "__main__":
    main()

