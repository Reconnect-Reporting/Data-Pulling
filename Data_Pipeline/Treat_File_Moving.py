import os
import time
import shutil
import random
from pathlib import Path
from typing import Iterable

# --- Where from / where to (auto-uses current username) ---
def get_downloads_dir() -> Path:
    return Path.home() / "Downloads"

def get_treat_dir() -> Path:
    return Path.home() / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Raw Data" / "Treat"

# --- Helpers ---
def _is_incomplete(p: Path) -> bool:
    # Skip in-progress Chrome files or obvious temps
    return p.suffix.lower() == ".crdownload" or p.name.lower().endswith(".tmp")

def _wait_size_stable(p: Path, checks: int = 3, delay: float = 0.4) -> bool:
    """Return True if file size is stable across a few checks (avoids racing an active download)."""
    try:
        last = p.stat().st_size
        for _ in range(checks):
            time.sleep(delay)
            cur = p.stat().st_size
            if cur != last:
                return False
            last = cur
        return True
    except FileNotFoundError:
        return False

def _atomic_replace_move(src: Path, dest: Path, retries: int = 6, backoff: float = 0.4):
    """
    Move src to dest directory as a temp file, then atomically replace dest with it.
    This overwrites existing files and works across volumes. Retries help with OneDrive locks.
    """
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.parent / (dest.name + ".~partial")

    # Clean any leftover temp
    try:
        if tmp.exists():
            tmp.unlink()
    except Exception:
        pass

    # Move source into destination folder as temp (handles cross-drive)
    shutil.move(str(src), str(tmp))

    # Atomically replace (overwrites). Retry if the target is briefly locked by OneDrive.
    for i in range(retries):
        try:
            os.replace(tmp, dest)  # atomic on same volume; overwrites if exists
            return
        except PermissionError:
            time.sleep(backoff * (i + 1))
        except Exception:
            raise
    # Final attempt:
    os.replace(tmp, dest)

def _gather_files(src_dir: Path, patterns: Iterable[str], recursive: bool) -> list[Path]:
    files: list[Path] = []
    if recursive:
        for pat in patterns:
            files.extend(p for p in src_dir.rglob(pat) if p.is_file())
    else:
        for pat in patterns:
            files.extend(p for p in src_dir.glob(pat) if p.is_file())
    # Deduplicate while preserving order
    seen, out = set(), []
    for p in files:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out

# --- Main movers ---
def move_downloads_to_treat(
    patterns=("*.csv", "rpt_*"),
    recursive: bool = False,
    special_prefixes: tuple[str, ...] = (
        "Assessments",
        "Demographics",
        "External_Documents_Report",
        "Notes_Test",
        "IPCC_Goal",
        "IPCC_Issue",
    ),
) -> list[Path]:
    src_dir = get_downloads_dir()
    dst_dir = get_treat_dir()
    dst_dir.mkdir(parents=True, exist_ok=True)

    special_lc = tuple(s.lower() for s in special_prefixes)
    candidates = _gather_files(src_dir, patterns, recursive)

    moved: list[Path] = []
    for f in candidates:
        if _is_incomplete(f) or not _wait_size_stable(f):
            continue

        if f.name.lower().startswith(special_lc):
            continue

        dest = dst_dir / f.name
        try:
            _atomic_replace_move(f, dest)
            moved.append(dest)
            print(f"Moved (replaced if existed): {f.name} -> {dest}")
        except Exception as e:
            print(f"Skipped {f.name}: {e}")

    if not moved:
        print("No general files moved (nothing matched or files were still downloading).")
    else:
        print(f"Done general move. {len(moved)} file(s) moved to: {dst_dir}")
    return moved

def move_special_reports(
    recursive: bool = False,
    special_prefixes: tuple[str, ...] = (
        "Assessments",
        "Demographics",
        "External_Documents_Report",
        "Notes_Test",
        "IPCC_Goal",
        "IPCC_Issue",
    ),
) -> dict[str, Path]:
    src_dir = get_downloads_dir()
    dst_dir = get_treat_dir()
    dst_dir.mkdir(parents=True, exist_ok=True)

    normalize_name = {
        "Assessments": "Assessment",
        "Demographics": "Demographics",
        "External_Documents_Report": "External_Documents_Report",
        "Notes_Test": "Notes_Test",
        "IPCC_Goal": "IPCC_Goal",
        "IPCC_Issue": "IPCC_Issue",
    }

    picked: dict[str, Path] = {}
    for prefix in special_prefixes:
        pool = (
            [p for p in src_dir.rglob(f"{prefix}*") if p.is_file()]
            if recursive
            else [p for p in src_dir.glob(f"{prefix}*") if p.is_file()]
        )

        pool = [p for p in pool if not _is_incomplete(p) and _wait_size_stable(p)]

        if not pool:
            print(f"No ready files found for prefix '{prefix}'.")
            continue

        choice = random.choice(pool)
        base = normalize_name.get(prefix, prefix)
        dest = dst_dir / f"{base}{choice.suffix}"

        try:
            _atomic_replace_move(choice, dest)
            picked[prefix] = dest
            print(f"Picked & moved '{choice.name}' for '{prefix}' -> {dest.name}")
        except Exception as e:
            print(f"Failed moving '{choice.name}' for '{prefix}': {e}")

    if not picked:
        print("No special reports moved.")
    else:
        print("Special reports moved:", ", ".join(f"{k} -> {v.name}" for k, v in picked.items()))
    return picked

def run(
    patterns=("*.csv", "rpt_*"),
    recursive: bool = False,
    special_prefixes: tuple[str, ...] = (
        "Assessments",
        "Demographics",
        "External_Documents_Report",
        "Notes_Test",
        "IPCC_Goal",
        "IPCC_Issue",
    ),
):
    moved_general = move_downloads_to_treat(
        patterns=patterns,
        recursive=recursive,
        special_prefixes=special_prefixes
    )
    moved_special = move_special_reports(
        recursive=recursive,
        special_prefixes=special_prefixes
    )
    return {"general": moved_general, "special": moved_special}

# Optional alias so _get_runner() can also use main()
def main():
    return run()

if __name__ == "__main__":
    main()
