# -*- coding: utf-8 -*-
import os
import re
from pathlib import Path
from datetime import datetime, timedelta
import win32com.client as win32

# ========= CONFIG =========
OUTLOOK_ACCOUNT    = "recreporting@reconnect.on.ca"   # Mailbox display name; if unsure, leave as-is (falls back to default store)
SENDER_EMAIL       = "no-reply@alayamail.com"         # Sender to filter on
INCLUDE_SUBFOLDERS = False                            # Set True to scan Inbox subfolders as well

# Base save directory (per your request)
USER = os.getlogin()
BASE_SAVE_DIR = rf"C:\Users\{USER}\OneDrive - Reconnect Community Health Services (1)\data\Raw Data\AlayaCare"

# ========= HELPERS =========
def ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)

def sanitize_filename(name: str) -> str:
    # Replace Windows-illegal characters
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()

def sanitize_folder(name: str) -> str:
    # Folders: also avoid trailing dots/spaces
    name = sanitize_filename(name)
    return name.rstrip(" .")

def outlook_dt_string(dt: datetime) -> str:
    # Outlook Restrict expects US format 12-hour clock
    return dt.strftime("%m/%d/%Y %I:%M %p")

def get_inbox_store_by_display(namespace, display_name: str):
    # Try to find a store whose DisplayName matches; else default store
    for store in namespace.Stores:
        try:
            if store.DisplayName.lower() == display_name.lower():
                return store
        except Exception:
            pass
    return namespace.Stores.Item(1)

def iter_folder_and_children(folder):
    yield folder
    for f in folder.Folders:
        yield from iter_folder_and_children(f)

def get_ext_from_attachment(att) -> str:
    fn = att.FileName or ""
    ext = os.path.splitext(fn)[1]
    return ext if ext else ".dat"

# --- Name parsing ---
# We look for the first YYYY-MM-DD in the filename, keep everything up to the end of that date,
# drop any trailing time like 'T1451' etc.
DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")

def split_type_and_date(filename_stem: str):
    """
    Returns (type_prefix, date_yyyy_mm_dd) or (None, None) if not found.
    Example: 'Meals_on_Wheels_2025-08-28T1451' -> ('Meals_on_Wheels', '2025-08-28')
    """
    m = DATE_RE.search(filename_stem)
    if not m:
        return None, None
    date_str = m.group(1)
    prefix = filename_stem[:m.start()].rstrip("_ -")
    prefix = prefix if prefix else "Unknown_Type"
    return prefix, date_str

def make_unique_path(path: Path) -> Path:
    """
    If path exists, append (_2), (_3), ... before extension.
    """
    if not path.exists():
        return path
    stem, ext = path.stem, path.suffix
    i = 2
    while True:
        candidate = path.with_name(f"{stem}_{i}{ext}")
        if not candidate.exists():
            return candidate
        i += 1

# ========= MAIN =========
def main():
    ensure_dir(BASE_SAVE_DIR)

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    store = get_inbox_store_by_display(outlook, OUTLOOK_ACCOUNT)
    inbox = store.GetDefaultFolder(6)  # 6 = olFolderInbox

    now = datetime.now()
    start_time = now - timedelta(days=7)

    sender_filter = f"[SenderEmailAddress] = '{SENDER_EMAIL}'"
    start_filter  = f"[ReceivedTime] >= '{outlook_dt_string(start_time)}'"
    end_filter    = f"[ReceivedTime] <  '{outlook_dt_string(now)}'"

    folders = [inbox] if not INCLUDE_SUBFOLDERS else list(iter_folder_and_children(inbox))

    total_saved = 0

    for f in folders:
        items = f.Items
        items.Sort("[ReceivedTime]", True)  # newest first
        try:
            items = items.Restrict(sender_filter).Restrict(start_filter).Restrict(end_filter)
        except Exception:
            # If Restrict fails for any reason, skip this folder
            continue

        # Iterate a concrete list to avoid COM moving-target issues
        for mail in list(items):
            try:
                if getattr(mail, "Class", None) != 43:  # 43 = MailItem
                    continue
                if mail.Attachments.Count == 0:
                    continue

                for i in range(1, mail.Attachments.Count + 1):
                    att = mail.Attachments.Item(i)
                    raw_name = (att.FileName or "").strip()
                    if not raw_name:
                        continue

                    # Split into stem/ext safely
                    stem, ext = os.path.splitext(raw_name)
                    if not ext:
                        ext = get_ext_from_attachment(att)

                    type_prefix, date_str = split_type_and_date(stem)

                    if type_prefix and date_str:
                        folder_name = sanitize_folder(type_prefix)
                        file_name = sanitize_filename(f"{type_prefix}_{date_str}{ext}")
                    else:
                        # Fallback if no date pattern found: bucket into 'Unknown_Type'
                        folder_name = "Unknown_Type"
                        file_name = sanitize_filename(raw_name)

                    dest_dir = Path(BASE_SAVE_DIR) / folder_name
                    ensure_dir(str(dest_dir))

                    dest_path = dest_dir / file_name
                    dest_path = make_unique_path(dest_path)

                    try:
                        att.SaveAsFile(str(dest_path))
                        print(f"Saved: {dest_path}")
                        total_saved += 1
                    except Exception as e:
                        print(f"Failed to save '{raw_name}': {e}")

            except Exception as e:
                print(f"Skipped one item due to error: {e}")

    print(f"Done. Total files saved: {total_saved}")

if __name__ == "__main__":
    main()
