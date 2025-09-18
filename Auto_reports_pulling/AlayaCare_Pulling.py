# -*- coding: utf-8 -*-
import os, re
from pathlib import Path
from datetime import datetime, timedelta
import win32com.client as win32

# ========= CONFIG =========
ACCOUNT_HINT  = "recreporting@reconnect.on.ca"  # mailbox SMTP or display name
SENDER_EMAIL  = "no-reply@alayamail.com"        # sender to match
DAYS_BACK     = 7                               
SUBFOLDERNAME = "Processed (AlayaCare)"         # must exist under Inbox

USER = os.getlogin()
BASE_SAVE_DIR = rf"C:\Users\{USER}\OneDrive - Reconnect Community Health Services (1)\data\Raw Data\AlayaCare"

# ========= HELPERS =========
def ensure_dir(p: str | Path):
    os.makedirs(p, exist_ok=True)

def sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", str(name)).strip()

def sanitize_folder(name: str) -> str:
    name = sanitize_filename(name)
    return name.rstrip(" .")

def get_ext_from_attachment(att) -> str:
    fn = att.FileName or ""
    ext = os.path.splitext(fn)[1]
    return ext if ext else ".dat"

DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")
def split_type_and_date(stem: str):
    m = DATE_RE.search(stem)
    if not m:
        return None, None
    date_str = m.group(1)
    prefix = stem[:m.start()].rstrip("_ -")
    return (prefix if prefix else "Unknown_Type"), date_str

def _fmt_outlook(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")  # Outlook Restrict format

def get_sender_smtp(mail) -> str:
    try:
        if getattr(mail, "SenderEmailType", "") == "EX":
            ex = mail.Sender.GetExchangeUser()
            if ex and ex.PrimarySmtpAddress:
                return ex.PrimarySmtpAddress.lower().strip()
    except Exception:
        pass
    try:
        return (mail.SenderEmailAddress or "").lower().strip()
    except Exception:
        return ""

def find_store(namespace, account_hint: str):
    """Try Accounts first (SMTP/display), then Stores by DisplayName; fallback to first store."""
    hint = (account_hint or "").lower().strip()
    try:
        for acct in namespace.Accounts:
            smtp = (getattr(acct, "SmtpAddress", "") or "").lower().strip()
            disp = (getattr(acct, "DisplayName", "") or "").lower().strip()
            if hint and (hint == smtp or hint == disp or hint in smtp or hint in disp):
                try:
                    return acct.DeliveryStore  # best match
                except Exception:
                    pass
    except Exception:
        pass

    try:
        for store in namespace.Stores:
            disp = (getattr(store, "DisplayName", "") or "").lower().strip()
            if hint and (hint == disp or hint in disp):
                return store
    except Exception:
        pass

    # fallback: first store
    return namespace.Stores.Item(1)

def get_inbox(store):
    # Per-store Inbox (not the default store’s Inbox)
    return store.GetDefaultFolder(6)  # olFolderInbox

def get_existing_subfolder(parent, name: str):
    try:
        return parent.Folders.Item(name)
    except Exception:
        return None

# ========= MAIN =========
def main():
    ensure_dir(BASE_SAVE_DIR)

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    store = find_store(ns, ACCOUNT_HINT)
    inbox = get_inbox(store)

    store_name = getattr(store, "DisplayName", "<unknown store>")
    print(f"Using store: {store_name}")
    print(f"Inbox folder name: {getattr(inbox, 'Name', 'Inbox')}")

    processed_folder = get_existing_subfolder(inbox, SUBFOLDERNAME)
    print(f"{'Found' if processed_folder else 'Did NOT find'} subfolder: {SUBFOLDERNAME}")

    now = datetime.now()
    start_time = now - timedelta(days=DAYS_BACK)

    # Date Restrict (sender filtered in Python for reliability)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    try:
        items = items.Restrict(f"[ReceivedTime] >= '{_fmt_outlook(start_time)}'") \
                     .Restrict(f"[ReceivedTime] < '{_fmt_outlook(now)}'")
    except Exception:
        pass

    target_sender = SENDER_EMAIL.lower().strip()
    total_saved = 0
    matched_msgs = 0

    for mail in list(items):
        try:
            if getattr(mail, "Class", None) != 43:  # MailItem
                continue

            if get_sender_smtp(mail) != target_sender:
                continue
            matched_msgs += 1

            if mail.Attachments.Count == 0:
                continue

            saved_any = False
            for i in range(1, mail.Attachments.Count + 1):
                att = mail.Attachments.Item(i)
                raw_name = (att.FileName or "").strip()
                if not raw_name:
                    continue

                stem, ext = os.path.splitext(raw_name)
                if not ext:
                    ext = get_ext_from_attachment(att)

                type_prefix, date_str = split_type_and_date(stem)
                if type_prefix and date_str:
                    folder_name = sanitize_folder(type_prefix)
                    file_name = sanitize_filename(f"{type_prefix}_{date_str}{ext}")
                else:
                    folder_name = "Unknown_Type"
                    file_name = sanitize_filename(raw_name)

                dest_dir = Path(BASE_SAVE_DIR) / folder_name
                ensure_dir(dest_dir)
                dest_path = dest_dir / file_name

                if dest_path.exists():
                    print(f"Skip (already exists): {dest_path}")
                    continue

                try:
                    att.SaveAsFile(str(dest_path))
                    print(f"Saved: {dest_path}")
                    total_saved += 1
                    saved_any = True
                except Exception as e:
                    print(f"Failed to save '{raw_name}': {e}")

            # Move to Inbox\Processed (AlayaCare) ONLY if it exists
            if saved_any and processed_folder is not None:
                try:
                    mail.Move(processed_folder)
                except Exception as e:
                    print(f"Warning: couldn't move to '{SUBFOLDERNAME}': {e}")

        except Exception as e:
            print(f"Skipped one item due to error: {e}")

    print(f"Done. Messages matched (sender={SENDER_EMAIL}): {matched_msgs}, files saved: {total_saved}")

if __name__ == "__main__":
    main()
