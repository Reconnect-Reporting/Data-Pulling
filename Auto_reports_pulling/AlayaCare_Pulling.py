# -*- coding: utf-8 -*-
import os
import re
from pathlib import Path
from datetime import datetime, timedelta
import win32com.client as win32

# ========= CONFIG =========
TARGET_STORE_DISPLAY = "recreporting@reconnect.on.ca"
SENDER_EMAIL         = "no-reply@alayamail.com"
DAYS_BACK            = 7

USER = os.getlogin()
BASE_SAVE_DIR = Path(
    rf"C:\Users\{USER}\OneDrive - Reconnect Community Health Services (1)\data\Raw Data\AlayaCare"
)

# ========= HELPERS =========
def ensure_dir(p: Path | str) -> None:
    Path(p).mkdir(parents=True, exist_ok=True)

def sanitize_filename(s: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", str(s)).strip()

def sanitize_folder(s: str) -> str:
    return sanitize_filename(s).rstrip(" .")

def _fmt_outlook(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")

def canon(s: str) -> str:
    return (s or "").replace("\u00A0", " ").strip().lower()

def get_ext_from_attachment(att) -> str:
    ext = os.path.splitext((att.FileName or ""))[1]
    return ext if ext else ".dat"

DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")

def split_type_and_date(stem: str):
    m = DATE_RE.search(stem)
    if not m:
        return None, None

    prefix = stem[:m.start()].rstrip("_ -") or "Unknown_Type"
    return prefix, m.group(1)

def get_sender_smtp(mail) -> str:
    try:
        if getattr(mail, "SenderEmailType", "") == "EX":
            ex = mail.Sender.GetExchangeUser()
            if ex and ex.PrimarySmtpAddress:
                return ex.PrimarySmtpAddress.lower().strip()
    except Exception:
        pass

    return (getattr(mail, "SenderEmailAddress", "") or "").lower().strip()

def pick_store_by_display(namespace, display_name: str):
    want = canon(display_name)

    for st in namespace.Stores:
        if canon(getattr(st, "DisplayName", "")) == want:
            return st

    available = ", ".join(
        (getattr(s, "DisplayName", "") or "").strip()
        for s in namespace.Stores
    )

    raise SystemExit(
        f"Mailbox '{display_name}' not found. Available mailboxes: {available}"
    )

def get_inbox(store):
    return store.GetDefaultFolder(6)  # 6 = olFolderInbox

def get_reports_folder(inbox):
    try:
        return inbox.Folders["Reports"]
    except Exception:
        print("Warning: Reports folder was not found under Inbox.")
        return None

def process_folder(folder, sender, start_time, now):
    print(f"\nScanning folder: {folder.Name}")

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    try:
        items = items.Restrict(f"[ReceivedTime] >= '{_fmt_outlook(start_time)}'")
        items = items.Restrict(f"[ReceivedTime] <  '{_fmt_outlook(now)}'")
    except Exception:
        print("Warning: Outlook Restrict failed. Proceeding without date filter.")

    total_saved = 0
    matched_msgs = 0

    for mail in list(items):
        try:
            if getattr(mail, "Class", None) != 43:
                continue

            if get_sender_smtp(mail) != sender:
                continue

            matched_msgs += 1
            atts = getattr(mail, "Attachments", None)

            if atts and atts.Count > 0:
                for i in range(1, atts.Count + 1):
                    att = atts.Item(i)
                    raw_name = (att.FileName or "").strip()

                    if not raw_name:
                        continue

                    stem, ext = os.path.splitext(raw_name)

                    if not ext:
                        ext = get_ext_from_attachment(att)

                    prefix, date_str = split_type_and_date(stem)

                    if prefix and date_str:
                        folder_name = sanitize_folder(prefix)
                        file_name = sanitize_filename(f"{prefix}_{date_str}{ext}")
                    else:
                        folder_name = "Unknown_Type"
                        file_name = sanitize_filename(raw_name)

                    dest_dir = BASE_SAVE_DIR / folder_name
                    ensure_dir(dest_dir)

                    dest_path = dest_dir / file_name

                    if dest_path.exists():
                        print(f"Skip existing file: {dest_path}")
                        continue

                    try:
                        att.SaveAsFile(str(dest_path))
                        print(f"Saved: {dest_path}")
                        total_saved += 1
                    except Exception as e:
                        print(f"Failed to save '{raw_name}': {e}")

        except Exception as e:
            print(f"Skipped one item due to error: {e}")

    print(
        f"Finished folder: {folder.Name} | "
        f"Matched emails: {matched_msgs}, attachments saved: {total_saved}"
    )

    return matched_msgs, total_saved

# ========= MAIN =========
def main():
    ensure_dir(BASE_SAVE_DIR)

    app = win32.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")

    store = pick_store_by_display(ns, TARGET_STORE_DISPLAY)
    inbox = get_inbox(store)
    reports = get_reports_folder(inbox)

    print(f"Mailbox found: {(getattr(store, 'DisplayName', '') or '').strip()}")

    now = datetime.now()
    start_time = now - timedelta(days=DAYS_BACK)

    sender = (SENDER_EMAIL or "").lower().strip()

    total_matched = 0
    total_saved = 0

    # Scan main Inbox
    matched, saved = process_folder(inbox, sender, start_time, now)
    total_matched += matched
    total_saved += saved

    # Scan Inbox > Reports
    if reports is not None:
        matched, saved = process_folder(reports, sender, start_time, now)
        total_matched += matched
        total_saved += saved

    print("\n========== DONE ==========")
    print(f"Total matched emails from {SENDER_EMAIL}: {total_matched}")
    print(f"Total attachments saved: {total_saved}")

if __name__ == "__main__":
    main()