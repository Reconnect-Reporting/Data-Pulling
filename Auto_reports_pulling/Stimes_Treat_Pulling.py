import os
import time
import calendar
from pathlib import Path
from datetime import datetime, date

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


# ===================== CONFIG =====================
LOGIN_URL = "https://sharedservices.treat.ca/treat/logon"
REPORTS_URL = "https://sharedservices.treat.ca/treat/app/reports?page=newReports"

ORG_ID = "stimes"
USERNAME = "yxu"
PASSWORD = "Jiazhuo2018#111"

DOWNLOAD_DIR = Path.home() / "Downloads"


# ===================== DATE HELPERS =====================
def get_yearly_cycle_dates():
    """
    Yearly cycle logic:
    Example: If today is May 11, 2026,
    Start Date = Jun 1, 2025
    End Date = May 11, 2026
    """
    today = datetime.today().date()

    if today.month == 12:
        start_dt = date(today.year, 1, 1)
    else:
        start_dt = date(today.year - 1, today.month + 1, 1)

    return start_dt.strftime("%d-%b-%Y"), today.strftime("%d-%b-%Y")


def subtract_months(original_date, months):
    month = original_date.month - months
    year = original_date.year

    while month <= 0:
        month += 12
        year -= 1

    last_day = calendar.monthrange(year, month)[1]
    day = min(original_date.day, last_day)

    return date(year, month, day)


def get_last_6_month_dates():
    """
    Six calendar months including the current month.
    Example: If today is Aug 31, 2026,
    Start Date = Mar 1, 2026
    End Date = Aug 31, 2026
    """
    today = datetime.today().date()

    month = today.month - 5
    year = today.year

    while month <= 0:
        month += 12
        year -= 1

    start_dt = date(year, month, 1)

    return start_dt.strftime("%d-%b-%Y"), today.strftime("%d-%b-%Y")


# ===================== DOWNLOAD HELPERS =====================
def is_temp_download_file(file_path):
    """
    Returns True if Chrome is still downloading or using a temporary file.
    """
    name = file_path.name.lower()
    return (
        name.endswith(".crdownload")
        or name.endswith(".tmp")
        or name.endswith(".download")
    )


def snapshot_downloads(download_dir):
    """
    Take a snapshot of current CSV files in the download folder.

    Only CSV files are tracked. This prevents unrelated browser files
    such as downloads.htm from being mistaken for a downloaded report.
    """
    download_dir = Path(download_dir)
    snapshot = {}

    if not download_dir.exists():
        download_dir.mkdir(parents=True, exist_ok=True)

    for file in download_dir.iterdir():
        if not file.is_file():
            continue

        if is_temp_download_file(file):
            continue

        # IMPORTANT: only track exported CSV report files
        if file.suffix.lower() != ".csv":
            continue

        try:
            snapshot[file.name] = (
                file.stat().st_mtime,
                file.stat().st_size
            )
        except FileNotFoundError:
            pass

    return snapshot


def wait_until_file_stable(file_path, stable_seconds=2, timeout=60):
    """
    Wait until a downloaded file stops changing size.
    This avoids renaming while the browser is still writing the file.
    """
    file_path = Path(file_path)
    end_time = time.time() + timeout

    last_size = None
    stable_start = None

    while time.time() < end_time:
        if not file_path.exists():
            time.sleep(0.5)
            continue

        try:
            current_size = file_path.stat().st_size
        except FileNotFoundError:
            time.sleep(0.5)
            continue

        if current_size == last_size:
            if stable_start is None:
                stable_start = time.time()

            if time.time() - stable_start >= stable_seconds:
                return file_path
        else:
            last_size = current_size
            stable_start = None

        time.sleep(0.5)

    raise TimeoutError(f"File did not become stable: {file_path}")


def wait_for_new_download(download_dir, before_snapshot, timeout=240):
    """
    Wait for a new or updated CSV report download.

    IMPORTANT:
    - Only CSV files are considered valid report downloads.
    - Unrelated files such as downloads.htm are completely ignored.
    - Chrome temporary files are ignored until the download completes.
    """
    download_dir = Path(download_dir)
    end_time = time.time() + timeout

    while time.time() < end_time:

        current_files = {}

        for file in download_dir.iterdir():
            if not file.is_file():
                continue

            # Ignore Chrome temporary files
            if is_temp_download_file(file):
                continue

            # IMPORTANT: only consider CSV report files
            if file.suffix.lower() != ".csv":
                continue

            try:
                current_files[file.name] = (
                    file.stat().st_mtime,
                    file.stat().st_size
                )
            except FileNotFoundError:
                pass

        changed_files = []

        for filename, file_info in current_files.items():

            # Brand-new CSV filename
            if filename not in before_snapshot:
                changed_files.append(download_dir / filename)

            # Existing filename that Chrome/Treat replaced or updated
            elif before_snapshot[filename] != file_info:
                changed_files.append(download_dir / filename)

        if changed_files:

            # Pick the newest changed CSV only
            newest_file = max(
                changed_files,
                key=lambda f: f.stat().st_mtime
            )

            print(f"Detected downloaded CSV: {newest_file.name}")

            # Make sure the actual CSV has finished writing
            wait_until_file_stable(
                newest_file,
                stable_seconds=2,
                timeout=60
            )

            return newest_file

        time.sleep(0.5)

    raise TimeoutError(
        "No new completed CSV download was detected within the timeout period."
    )


def add_stimes_to_filename(file_path, forced_base_name=None):
    """
    Rename downloaded file by adding _Stimes.

    Examples:
        ADTCensusCAN.csv
        -> ADTCensusCAN_Stimes.csv

    If that already exists:
        -> ADTCensusCAN_Stimes_1.csv
        -> ADTCensusCAN_Stimes_2.csv
        -> etc.

    Existing files are NEVER overwritten.
    """

    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(
            f"Downloaded file not found: {file_path}"
        )

    if is_temp_download_file(file_path):
        raise RuntimeError(
            f"Refusing to rename temporary download file: {file_path}"
        )

    # Keep original extension
    suffix = file_path.suffix if file_path.suffix else ".csv"

    # Use forced filename if provided
    if forced_base_name:
        base_name = forced_base_name
    else:
        base_name = file_path.stem

    # Prevent duplicate _Stimes
    if base_name.lower().endswith("_stimes"):
        base_name = base_name[:-7]

    # Desired filename
    new_path = file_path.parent / f"{base_name}_Stimes{suffix}"

    # =====================================================
    # DO NOT OVERWRITE EXISTING FILES
    # =====================================================
    counter = 1

    while new_path.exists():
        new_path = (
            file_path.parent /
            f"{base_name}_Stimes_{counter}{suffix}"
        )
        counter += 1

    # Rename downloaded file
    file_path.rename(new_path)

    print("Renamed downloaded file:")
    print(f"  From: {file_path.name}")
    print(f"  To:   {new_path.name}")

    return new_path

# ===================== DRIVER SETUP =====================
def build_driver():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    opts = Options()
    opts.add_argument("--start-maximized")

    opts.set_capability("acceptInsecureCerts", True)
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--ignore-ssl-errors=yes")

    opts.add_argument("--log-level=3")
    opts.add_argument("--disable-logging")
    opts.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])

    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }

    opts.add_experimental_option("prefs", prefs)

    service = Service(log_output=os.devnull)
    driver = webdriver.Chrome(service=service, options=opts)

    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": str(DOWNLOAD_DIR)
            }
        )
    except Exception:
        pass

    return driver


# ===================== COMMON HELPERS =====================
def bypass_ssl_page_if_needed(driver):
    try:
        details_btn = driver.find_elements(By.ID, "details-button")
        if details_btn:
            details_btn[0].click()
            time.sleep(1)

            proceed_link = driver.find_elements(By.ID, "proceed-link")
            if proceed_link:
                proceed_link[0].click()
                time.sleep(2)
    except Exception:
        pass


def safe_type(driver, wait, by, value, text, field_name="field"):
    """
    Wait until a field is clickable, then type into it.
    Falls back to JavaScript if Selenium cannot interact normally.
    """
    element = wait.until(
        EC.element_to_be_clickable((by, value))
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    time.sleep(0.5)

    try:
        element.click()
        element.clear()
        element.send_keys(text)
    except Exception:
        driver.execute_script("arguments[0].value = '';", element)
        driver.execute_script("arguments[0].value = arguments[1];", element, text)

    print(f"Entered {field_name}.")


def enter_dates(driver, wait, start_date, end_date):
    print("Start Date:", start_date)
    print("End Date:", end_date)

    start_box = wait.until(
        EC.element_to_be_clickable((By.ID, "start"))
    )

    driver.execute_script("""
        arguments[0].removeAttribute('readonly');
        arguments[0].readOnly = false;
    """, start_box)

    start_box.click()
    start_box.clear()
    start_box.send_keys(start_date)

    end_box = wait.until(
        EC.element_to_be_clickable((By.ID, "end"))
    )

    driver.execute_script("""
        arguments[0].removeAttribute('readonly');
        arguments[0].readOnly = false;
    """, end_box)

    end_box.click()
    end_box.clear()
    end_box.send_keys(end_date)
    end_box.send_keys(Keys.ESCAPE)

    print("Dates entered.")


def click_generate_report(driver, wait, report_name=""):
    generate_btn = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//input[@type='button' and @value='Generate Report']"
            )
        )
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", generate_btn)
    time.sleep(0.5)

    try:
        generate_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", generate_btn)

    if report_name:
        print(f"Clicked Generate Report for {report_name}.")
    else:
        print("Clicked Generate Report.")


def export_csv(driver, report_name="", forced_base_name=None):
    """
    Exports CSV, waits for the download, then renames it by adding _Stimes.
    If forced_base_name is provided, the final file name will use that base name.
    """
    driver.switch_to.default_content()

    before_snapshot = snapshot_downloads(DOWNLOAD_DIR)

    WebDriverWait(driver, 300).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "reportViewer"))
    )

    print("Switched into reportViewer iframe.")

    export_dropdown = WebDriverWait(driver, 300).until(
        EC.presence_of_element_located((By.ID, "rvViewer_ctl01_ctl05_ctl00"))
    )

    Select(export_dropdown).select_by_visible_text("CSV (comma delimited)")
    print("Selected CSV format.")

    export_btn = WebDriverWait(driver, 300).until(
        EC.element_to_be_clickable((By.ID, "rvViewer_ctl01_ctl05_ctl01"))
    )

    try:
        export_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", export_btn)

    if report_name:
        print(f"Clicked Export for {report_name}.")
    else:
        print("Clicked Export.")

    driver.switch_to.default_content()

    downloaded_file = wait_for_new_download(DOWNLOAD_DIR, before_snapshot, timeout=240)

    renamed_file = add_stimes_to_filename(
        downloaded_file,
        forced_base_name=forced_base_name
    )

    print(f"Download completed and renamed: {renamed_file}")

    time.sleep(2)


def go_to_main_reports_page(driver, wait, wait_for_report_id=None):
    driver.switch_to.default_content()
    driver.get(REPORTS_URL)

    if wait_for_report_id:
        wait.until(
            EC.presence_of_element_located((By.ID, wait_for_report_id))
        )

    print("Back to main reports page.")
    print("Current URL:", driver.current_url)


def open_report(driver, wait, report_id, report_name):
    report_el = wait.until(
        EC.element_to_be_clickable((By.ID, report_id))
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", report_el)
    time.sleep(0.5)

    try:
        report_el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", report_el)

    print(f"Clicked {report_name}.")
    print("Current URL:", driver.current_url)


def choose_all_programs(driver, wait):
    program_div = wait.until(
        EC.element_to_be_clickable((By.ID, "programDiv"))
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", program_div)
    time.sleep(0.5)

    try:
        program_div.click()
    except Exception:
        driver.execute_script("arguments[0].click();", program_div)

    print("Opened Program options.")

    wait.until(
        EC.visibility_of_element_located((By.ID, "programCheckBoxDiv"))
    )

    all_program_checkbox = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//input[@type='checkbox' and contains(@onclick, \"toggleAllOptions(this.checked, 'program', '2')\")]"
            )
        )
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", all_program_checkbox)
    time.sleep(0.3)

    if not all_program_checkbox.is_selected():
        try:
            all_program_checkbox.click()
        except Exception:
            driver.execute_script("arguments[0].click();", all_program_checkbox)

    print("Selected all programs.")


def run_report(driver, wait, report_id, report_name, date_mode="yearly", select_programs=False, forced_base_name=None):
    """
    date_mode:
      - "yearly" = yearly cycle date logic
      - "last_6_months" = last 6 months from today
    """
    go_to_main_reports_page(driver, wait, wait_for_report_id=report_id)
    open_report(driver, wait, report_id, report_name)

    if date_mode == "last_6_months":
        start_date, end_date = get_last_6_month_dates()
    else:
        start_date, end_date = get_yearly_cycle_dates()

    enter_dates(driver, wait, start_date, end_date)

    if select_programs:
        choose_all_programs(driver, wait)

    click_generate_report(driver, wait, report_name)

    export_csv(
        driver,
        report_name=report_name,
        forced_base_name=forced_base_name
    )


# ===================== FRAME / DRILLTHROUGH HELPERS =====================
def switch_into_frame_with_xpath(driver, xpath_expr, timeout=120, max_depth=8):
    """
    Search nested frames until an XPath element exists.
    Leaves driver focused inside the frame containing the element.
    """
    end = time.time() + timeout

    while time.time() < end:
        driver.switch_to.default_content()

        def dfs(depth):
            try:
                elements = driver.find_elements(By.XPATH, xpath_expr)
                if elements:
                    return True
            except Exception:
                pass

            if depth >= max_depth:
                return False

            frames = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")

            for i in range(len(frames)):
                try:
                    driver.switch_to.frame(i)
                except Exception:
                    continue

                if dfs(depth + 1):
                    return True

                driver.switch_to.parent_frame()

            return False

        if dfs(0):
            return True

        time.sleep(0.5)

    return False


def click_drillthrough_by_fragment(driver, fragment, report_name="Referral Summary Report"):
    """
    Clicks a drillthrough link by fragment, such as:
    - 120iT2:0
    - 132iT2:0

    If there are multiple matching links, this chooses the lowest/bottom-most visible one.
    """
    xpath = f"//a[contains(@onclick, \"Drillthrough','{fragment}'\")]"

    found = switch_into_frame_with_xpath(driver, xpath, timeout=180, max_depth=10)

    if not found:
        raise RuntimeError(f"Could not find drillthrough link for fragment: {fragment}")

    links = driver.find_elements(By.XPATH, xpath)

    visible_links = []

    for link in links:
        try:
            if link.is_displayed():
                rect = driver.execute_script("return arguments[0].getBoundingClientRect();", link)
                visible_links.append((link, rect["top"], rect["left"]))
        except Exception:
            pass

    if not visible_links:
        raise RuntimeError(f"Found drillthrough link for {fragment}, but none were visible.")

    target_link = sorted(visible_links, key=lambda x: (x[1], x[2]))[-1][0]

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_link)
    time.sleep(0.5)

    try:
        target_link.click()
    except Exception:
        driver.execute_script("arguments[0].click();", target_link)

    print(f"Clicked drillthrough {fragment} for {report_name}.")

    driver.switch_to.default_content()
    time.sleep(3)


def run_referral_summary_drillthrough(driver, wait, fragment, export_label):
    """
    Opens Referral Summary Report, enters yearly-cycle dates,
    selects all programs, generates report, clicks drillthrough,
    then exports CSV.
    """
    report_name = f"Referral Summary Report - {export_label}"

    go_to_main_reports_page(driver, wait, wait_for_report_id="report25B")
    open_report(driver, wait, "report25B", "Referral Summary Report")

    start_date, end_date = get_last_6_month_dates()
    enter_dates(driver, wait, start_date, end_date)

    choose_all_programs(driver, wait)

    click_generate_report(driver, wait, "Referral Summary Report")

    click_drillthrough_by_fragment(driver, fragment, "Referral Summary Report")

    export_csv(driver, report_name)


# ===================== LOGIN =====================
def login(driver, wait):
    driver.get(LOGIN_URL)
    time.sleep(2)

    bypass_ssl_page_if_needed(driver)

    safe_type(
        driver=driver,
        wait=wait,
        by=By.ID,
        value="orgUserID",
        text=ORG_ID,
        field_name="Organization ID"
    )

    continue_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "btnContinue"))
    )

    continue_btn.click()

    print("Clicked Continue.")

    safe_type(
        driver=driver,
        wait=wait,
        by=By.ID,
        value="username",
        text=USERNAME,
        field_name="User ID"
    )

    safe_type(
        driver=driver,
        wait=wait,
        by=By.ID,
        value="password",
        text=PASSWORD,
        field_name="Password"
    )

    logon_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "btnLogOn_TreatAuth"))
    )

    logon_btn.click()

    print("Clicked Log On.")

    time.sleep(3)


# ===================== MAIN =====================
def main():
    driver = build_driver()
    wait = WebDriverWait(driver, 30)

    try:
        login(driver, wait)

        go_to_main_reports_page(driver, wait, wait_for_report_id="report4B")

        # Report 1: ADT - Census Report
        run_report(
            driver=driver,
            wait=wait,
            report_id="report4B",
            report_name="ADT - Census Report",
            date_mode="last_6_months",
            select_programs=False,
            forced_base_name="ADTCensusCAN"
        )

        # Report 2: MIS - Statistics Report
        run_report(
            driver=driver,
            wait=wait,
            report_id="report13B",
            report_name="MIS - Statistics Report",
            date_mode="last_6_months",
            select_programs=True
        )

        # Report 3: Progress Notes - Staff Log for Finance
        run_report(
            driver=driver,
            wait=wait,
            report_id="report24B",
            report_name="Progress Notes - Staff Log for Finance",
            date_mode="last_6_months",
            select_programs=True
        )

        # Report 4: Workload - Proofing Report by Clinician
        run_report(
            driver=driver,
            wait=wait,
            report_id="report31B",
            report_name="Workload - Proofing Report by Clinician",
            date_mode="last_6_months",
            select_programs=False
        )

        # Report 5A: Referral Summary - New Referrals Drillthrough
        run_referral_summary_drillthrough(
            driver=driver,
            wait=wait,
            fragment="120iT2:0",
            export_label="New Referrals"
        )

        # Report 5B: Referral Summary - Waitlisted Drillthrough
        run_referral_summary_drillthrough(
            driver=driver,
            wait=wait,
            fragment="132iT2:0",
            export_label="Waitlisted"
        )

        print("All reports completed.")

    finally:
        time.sleep(5)
        driver.quit()


if __name__ == "__main__":
    main()