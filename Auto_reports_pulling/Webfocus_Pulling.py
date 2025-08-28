# pip install selenium webdriver-manager

import os
import time
from datetime import date
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager


# ----------------------- CONFIG -----------------------
URL_LOGIN = "https://www.caredove.com/user/account"
URL_REPORTS_MAIN = "https://www.caredove.com/organization/reports/orgid/3866"  # Reconnect Health Services
URL_REPORTS_WEST = "https://www.caredove.com/organization/reports/orgid/3851"  # West Community Connect Hub

EMAIL = os.getenv("CAREDOVE_EMAIL", "").strip()
PASSWORD = os.getenv("CAREDOVE_PASSWORD", "")

if not EMAIL or not PASSWORD:
    raise RuntimeError("Caredove credentials missing. Set CAREDOVE_EMAIL and CAREDOVE_PASSWORD in environment.")

DOWNLOAD_DIR = str(Path.home() / "Downloads" / "caredove")
DOWNLOAD_TIMEOUT = 120  # seconds


# ----------------------- DATE LOGIC (Sep 1 fiscal) -----------------------

def calendar_year_window():
    today = date.today()
    # next month
    next_month = today.month + 1
    year = today.year
    if next_month == 13:
        next_month = 1
        year += 1
    
    # start = first day of "next month" but in the previous year
    start = date(year - 1, next_month, 1)
    return start.strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d")



# ----------------------- HELPERS -----------------------
def wait_ready(driver, timeout=25):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def safe_click(driver, elem):
    try:
        elem.click(); return True
    except WebDriverException:
        pass
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(0.2); elem.click(); return True
    except WebDriverException:
        pass
    try:
        ActionChains(driver).move_to_element(elem).pause(0.1).click().perform(); return True
    except WebDriverException:
        pass
    try:
        driver.execute_script("arguments[0].click();", elem); return True
    except WebDriverException:
        return False

def snapshot_dir(download_dir: str):
    return set(os.listdir(download_dir))

def _size_stable(path: Path, checks=3, interval=0.5):
    try:
        prev = path.stat().st_size
        for _ in range(checks):
            time.sleep(interval)
            cur = path.stat().st_size
            if cur != prev:
                prev = cur
                return False
        return True
    except FileNotFoundError:
        return False

def wait_for_new_file(download_dir: str, before_snapshot: set, timeout: int) -> Path:
    start = time.time()
    d = Path(download_dir)
    while time.time() - start < timeout:
        after = set(os.listdir(download_dir))
        new_names = after - before_snapshot
        if new_names:
            candidates = [d / n for n in new_names if not str(n).endswith(".crdownload")]
            if candidates:
                newest = max(candidates, key=lambda p: p.stat().st_mtime)
                if _size_stable(newest):
                    return newest
        time.sleep(0.5)
    raise FileNotFoundError("No completed download detected within timeout")

def download_and_rename(driver, wait, download_btn_xpath: str, download_dir: str, new_filename: str):
    before = snapshot_dir(download_dir)
    download_btn = wait.until(EC.element_to_be_clickable((By.XPATH, download_btn_xpath)))
    if not safe_click(driver, download_btn):
        raise RuntimeError("Could not click download button")
    downloaded = wait_for_new_file(download_dir, before, timeout=DOWNLOAD_TIMEOUT)
    target = Path(download_dir) / f"{new_filename}.xlsx"
    if target.exists():
        target.unlink()
    downloaded.rename(target)
    print(f"Saved → {target}")

def open_report_dropdown(wait):
    return wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[@data-id='report_select' and contains(@class,'dropdown-toggle')]"
    )))

def choose_dropdown_option(wait, driver, option_text: str):
    btn = open_report_dropdown(wait)
    safe_click(driver, btn)
    option = wait.until(EC.element_to_be_clickable((
        By.XPATH, f"//span[@class='text' and normalize-space()='{option_text}']/.."
    )))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", option)
    if not safe_click(driver, option):
        raise RuntimeError(f"Could not click option: {option_text}")

def set_custom_date_range(wait, driver):
    """Ensure 'Custom' is selected and set start/end inputs using Sep 1 fiscal window."""
    start_str, end_str = calendar_year_window()

    # 1) Prefer clicking label[for='op6'] (some UIs hide the radio input)
    try:
        label = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='op6']")))
        safe_click(driver, label)
    except TimeoutException:
        pass  # fall through

    # 2) Also ensure the actual radio is checked (JS fallback)
    try:
        custom_radio = driver.find_element(By.ID, "op6")
        driver.execute_script(
            "if(arguments[0]){arguments[0].checked=true;"
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('click',{bubbles:true}));}", custom_radio
        )
    except NoSuchElementException:
        raise RuntimeError("CUSTOM date range radio (#op6) not found")

    # 3) Wait until date inputs are visible/enabled
    # Try by explicit IDs; if not, fall back to class 'rangeDates'
    start_input = None
    end_input = None
    try:
        start_input = wait.until(EC.visibility_of_element_located((By.ID, "startDate")))
    except TimeoutException:
        pass
    try:
        end_input = wait.until(EC.visibility_of_element_located((By.ID, "endDate")))
    except TimeoutException:
        pass

    if start_input is None or end_input is None:
        inputs = wait.until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "input.rangeDates")
        ))
        # choose first as start, last as end (common pattern)
        if inputs:
            if start_input is None:
                start_input = inputs[0]
            if end_input is None:
                end_input = inputs[-1] if len(inputs) > 1 else inputs[0]

    if start_input is None or end_input is None:
        raise RuntimeError("Could not locate start/end date inputs")

    # 4) Remove readonly/disabled if present, then set values via JS and fire events
    for inp, val in ((start_input, start_str), (end_input, end_str)):
        driver.execute_script(
            "arguments[0].removeAttribute('readonly');"
            "arguments[0].removeAttribute('disabled');", inp
        )
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
            inp, val
        )
    print(f"Set date range: {start_str} → {end_str}")


# ----------------------- MAIN FLOW -----------------------
def run_and_download():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    opts = Options()
    opts.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    wait = WebDriverWait(driver, 30)

    # --- Login ---
    driver.get(URL_LOGIN)
    wait.until(EC.presence_of_element_located((By.ID, "LoginForm_u_email"))).send_keys(EMAIL)
    pwd = wait.until(EC.presence_of_element_located((By.ID, "LoginForm_u_password")))
    pwd.clear(); pwd.send_keys(PASSWORD)
    wait.until(EC.element_to_be_clickable((By.ID, "formSubmit"))).click()
    wait_ready(driver)

    # --- Go to Reconnect Reports (orgid=3866) ---
    try:
        link = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='/organization/reports/orgid/3866']")))
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/organization/reports/orgid/3866']")))
        if not safe_click(driver, link):
            raise WebDriverException("Couldn't click reports link")
    except (TimeoutException, WebDriverException):
        driver.get(URL_REPORTS_MAIN)
    wait_ready(driver)

    # --- Reconnect: Inbound ---
    choose_dropdown_option(wait, driver, "Spreadsheet: Inbound Referral Raw Data")
    set_custom_date_range(wait, driver)
    download_and_rename(
        driver, wait,
        "//button[contains(@class,'download_report') and contains(., 'Download Data (.xlsx)')]",
        DOWNLOAD_DIR, "CaredoveBookingsReport_Reconnect"
    )

    # --- Reconnect: Outbound ---
    choose_dropdown_option(wait, driver, "Spreadsheet: Outbound Referral Raw Data")
    set_custom_date_range(wait, driver)
    download_and_rename(
        driver, wait,
        "//button[contains(@class,'download_report') and contains(., 'Download Data (.xlsx)')]",
        DOWNLOAD_DIR, "CaredoveBookingsReport_Outbound_Reconnect"
    )

    # --- Navigate directly to West Hub (orgid=3851) ---
    driver.get(URL_REPORTS_WEST)
    wait_ready(driver)

    # --- West: Inbound ---
    choose_dropdown_option(wait, driver, "Spreadsheet: Inbound Referral Raw Data")
    set_custom_date_range(wait, driver)
    download_and_rename(
        driver, wait,
        "//button[contains(@class,'download_report') and contains(., 'Download Data (.xlsx)')]",
        DOWNLOAD_DIR, "CaredoveBookingsReport_West"
    )

    # --- West: Outbound ---
    choose_dropdown_option(wait, driver, "Spreadsheet: Outbound Referral Raw Data")
    set_custom_date_range(wait, driver)
    download_and_rename(
        driver, wait,
        "//button[contains(@class,'download_report') and contains(., 'Download Data (.xlsx)')]",
        DOWNLOAD_DIR, "CaredoveBookingsReport_Outbound_West"
    )

    print("All four reports downloaded and renamed ✅")
    driver.quit()
    return

def run():
    return run_and_download()


if __name__ == "__main__":
    run_and_download()
