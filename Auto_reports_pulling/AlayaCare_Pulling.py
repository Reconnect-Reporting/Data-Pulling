# -*- coding: utf-8 -*-

import os
import re
from pathlib import Path
from datetime import datetime, timedelta

import win32com.client as win32


# ========= CONFIG =========

TARGET_STORE_DISPLAY = "recreporting@reconnect.on.ca"
SENDER_EMAIL = "no-reply@alayamail.com"
DAYS_BACK = 7

USER = os.getlogin()

BASE_SAVE_DIR = Path(
    rf"C:\Users\{USER}\OneDrive - Reconnect Community Health Services (1)\data\Raw Data\AlayaCare"
)

# New report
NEW_REPORT_NAME = "ADT_with_Zone_and_Reference_New_Referral_Date"


# ========= HELPERS =========

def ensure_dir(p: Path | str) -> None:
    Path(p).mkdir(parents=True, exist_ok=True)


def sanitize_filename(s: str) -> str:
    """
    Replace characters that are invalid in Windows filenames.
    """
    return re.sub(r'[<>:"/\\|?*]', "_", str(s)).strip()


def sanitize_folder(s: str) -> str:
    """
    Clean folder name and remove trailing spaces / periods.
    """
    return sanitize_filename(s).rstrip(" .")


def _fmt_outlook(dt: datetime) -> str:
    """
    Format datetime for Outlook Restrict.
    """
    return dt.strftime("%m/%d/%Y %I:%M %p")


def canon(s: str) -> str:
    """
    Normalize strings for mailbox name comparison.
    """
    return (
        (s or "")
        .replace("\u00A0", " ")
        .strip()
        .lower()
    )


def get_ext_from_attachment(att) -> str:
    """
    Get file extension from Outlook attachment.
    """
    ext = os.path.splitext(
        (att.FileName or "")
    )[1]

    return ext if ext else ".dat"


# Match dates like 2026-09-25
DATE_RE = re.compile(
    r"(\d{4}-\d{2}-\d{2})"
)


def split_type_and_date(stem: str):
    """
    Example:

    ADT_with_Zone_and_Reference_New_Referral_Date_2026-09-25

    returns:

    prefix:
        ADT_with_Zone_and_Reference_New_Referral_Date

    date:
        2026-09-25
    """

    match = DATE_RE.search(stem)

    if not match:
        return None, None

    prefix = (
        stem[:match.start()]
        .rstrip("_ -")
        or "Unknown_Type"
    )

    date_str = match.group(1)

    return prefix, date_str


def get_sender_smtp(mail) -> str:
    """
    Get actual SMTP sender address.
    Handles both Exchange and normal SMTP senders.
    """

    try:

        if getattr(mail, "SenderEmailType", "") == "EX":

            ex_user = mail.Sender.GetExchangeUser()

            if ex_user and ex_user.PrimarySmtpAddress:

                return (
                    ex_user.PrimarySmtpAddress
                    .lower()
                    .strip()
                )

    except Exception:
        pass

    return (
        getattr(
            mail,
            "SenderEmailAddress",
            ""
        )
        or ""
    ).lower().strip()


def pick_store_by_display(namespace, display_name: str):
    """
    Find Outlook mailbox/store by display name.
    """

    wanted = canon(display_name)

    for store in namespace.Stores:

        store_name = canon(
            getattr(
                store,
                "DisplayName",
                ""
            )
        )

        if store_name == wanted:
            return store

    available = ", ".join(
        (
            getattr(
                store,
                "DisplayName",
                ""
            )
            or ""
        ).strip()
        for store in namespace.Stores
    )

    raise SystemExit(
        f"Mailbox '{display_name}' not found. "
        f"Available mailboxes: {available}"
    )


def get_inbox(store):
    """
    Outlook folder 6 = Inbox
    """
    return store.GetDefaultFolder(6)


def get_reports_folder(inbox):
    """
    Try to find Inbox > Reports.
    """

    try:
        return inbox.Folders["Reports"]

    except Exception:

        print(
            "Warning: Reports folder was not found under Inbox."
        )

        return None


# ========= PROCESS OUTLOOK FOLDER =========

def process_folder(
    folder,
    sender,
    start_time,
    now
):

    print()
    print(
        f"Scanning folder: {folder.Name}"
    )

    items = folder.Items

    # Newest email first
    items.Sort(
        "[ReceivedTime]",
        True
    )

    # Try to filter to recent emails
    try:

        items = items.Restrict(
            f"[ReceivedTime] >= '{_fmt_outlook(start_time)}'"
        )

        items = items.Restrict(
            f"[ReceivedTime] < '{_fmt_outlook(now)}'"
        )

    except Exception as e:

        print(
            "Warning: Outlook Restrict failed. "
            "Proceeding without date filter."
        )

        print(
            f"Restrict error: {e}"
        )

    total_saved = 0
    matched_msgs = 0

    for mail in list(items):

        try:

            # 43 = Outlook MailItem
            if getattr(mail, "Class", None) != 43:
                continue

            actual_sender = get_sender_smtp(mail)

            # Only process AlayaCare reports
            if actual_sender != sender:
                continue

            matched_msgs += 1

            subject = (
                getattr(
                    mail,
                    "Subject",
                    ""
                )
                or ""
            )

            print()
            print(
                f"Matched email: {subject}"
            )

            print(
                f"Sender: {actual_sender}"
            )

            attachments = getattr(
                mail,
                "Attachments",
                None
            )

            if not attachments:
                continue

            if attachments.Count <= 0:
                continue

            for i in range(
                1,
                attachments.Count + 1
            ):

                att = attachments.Item(i)

                raw_name = (
                    att.FileName
                    or ""
                ).strip()

                if not raw_name:
                    continue

                print(
                    f"Attachment found: {raw_name}"
                )

                stem, ext = os.path.splitext(
                    raw_name
                )

                if not ext:

                    ext = get_ext_from_attachment(
                        att
                    )

                prefix, date_str = split_type_and_date(
                    stem
                )

                # ==========================================
                # Normal dated AlayaCare report
                # ==========================================

                if prefix and date_str:

                    folder_name = sanitize_folder(
                        prefix
                    )

                    file_name = sanitize_filename(
                        f"{prefix}_{date_str}{ext}"
                    )

                # ==========================================
                # Attachment doesn't contain YYYY-MM-DD
                # ==========================================

                else:

                    folder_name = "Unknown_Type"

                    file_name = sanitize_filename(
                        raw_name
                    )

                dest_dir = (
                    BASE_SAVE_DIR
                    / folder_name
                )

                # Automatically create report folder
                ensure_dir(
                    dest_dir
                )

                dest_path = (
                    dest_dir
                    / file_name
                )

                # Don't overwrite existing report
                if dest_path.exists():

                    print(
                        f"Skip existing file: {dest_path}"
                    )

                    continue

                try:

                    att.SaveAsFile(
                        str(dest_path)
                    )

                    print(
                        f"Saved: {dest_path}"
                    )

                    total_saved += 1

                except Exception as e:

                    print(
                        f"Failed to save "
                        f"'{raw_name}': {e}"
                    )

        except Exception as e:

            print(
                f"Skipped one item due to error: {e}"
            )

    print()
    print(
        f"Finished folder: {folder.Name}"
    )

    print(
        f"Matched emails: {matched_msgs}"
    )

    print(
        f"Attachments saved: {total_saved}"
    )

    return matched_msgs, total_saved


# ========= MAIN =========

def main():

    print(
        "========== ALAYACARE REPORT PULLING =========="
    )

    # Make sure root Raw Data folder exists
    ensure_dir(
        BASE_SAVE_DIR
    )

    # Explicitly make sure the NEW report folder exists
    new_report_folder = (
        BASE_SAVE_DIR
        / NEW_REPORT_NAME
    )

    ensure_dir(
        new_report_folder
    )

    print(
        f"Raw data root: {BASE_SAVE_DIR}"
    )

    print(
        f"New report folder ready: {new_report_folder}"
    )

    # Connect to Outlook desktop application
    app = win32.Dispatch(
        "Outlook.Application"
    )

    namespace = app.GetNamespace(
        "MAPI"
    )

    # Find shared mailbox
    store = pick_store_by_display(
        namespace,
        TARGET_STORE_DISPLAY
    )

    mailbox_name = (
        getattr(
            store,
            "DisplayName",
            ""
        )
        or ""
    ).strip()

    print(
        f"Mailbox found: {mailbox_name}"
    )

    inbox = get_inbox(
        store
    )

    reports_folder = get_reports_folder(
        inbox
    )

    now = datetime.now()

    start_time = (
        now
        - timedelta(
            days=DAYS_BACK
        )
    )

    sender = (
        SENDER_EMAIL
        or ""
    ).lower().strip()

    print(
        f"Looking back {DAYS_BACK} days"
    )

    print(
        f"Target sender: {sender}"
    )

    total_matched = 0
    total_saved = 0

    # ==========================================
    # Scan main Inbox
    # ==========================================

    matched, saved = process_folder(
        inbox,
        sender,
        start_time,
        now
    )

    total_matched += matched
    total_saved += saved

    # ==========================================
    # Scan Inbox > Reports
    # ==========================================

    if reports_folder is not None:

        matched, saved = process_folder(
            reports_folder,
            sender,
            start_time,
            now
        )

        total_matched += matched
        total_saved += saved

    # ==========================================
    # Summary
    # ==========================================

    print()
    print(
        "========== DONE =========="
    )

    print(
        f"Total matched emails from "
        f"{SENDER_EMAIL}: "
        f"{total_matched}"
    )

    print(
        f"Total attachments saved: "
        f"{total_saved}"
    )

    print()
    print(
        "Expected new report folder:"
    )

    print(
        new_report_folder
    )


if __name__ == "__main__":
    main()