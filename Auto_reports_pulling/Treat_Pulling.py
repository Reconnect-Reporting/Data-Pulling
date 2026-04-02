# -*- coding: utf-8 -*-
"""
Treat login + Synthesis reports downloader
- Bypasses SSL interstitial
- Runs multiple Synthesis reports and exports CSVs
- Handles WebFOCUS export to CSV
"""

import os
import time
from datetime import datetime, date
from pathlib import Path
from selenium import webdriver
import threading
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    MoveTargetOutOfBoundsException,
    WebDriverException,
)

# ===================== CONFIG =====================
LOGIN_URL = "https://sharedservices.treat.ca/treat/logon"
ORG_ID = "reconnect-local"
_download = Path.home() / "Downloads"
if not _download.exists():
    for var in ("OneDriveCommercial", "OneDrive"):
        d = os.environ.get(var)
        if d and (Path(d) / "Downloads").exists():
            _download = Path(d) / "Downloads"
            break

DOWNLOAD_DIR = _download

# Prefer env vars; falls back to literals if not set
USERNAME = os.getenv("TREAT_USERNAME", "yxu")
PASSWORD = os.getenv("TREAT_PASSWORD", "Jiazhuo2018#444")

# Date inputs
TODAY = datetime.today().date().strftime("%d-%b-%Y")
_today = datetime.today().date()

if _today.month == 12:
    # December → Jan of SAME year
    _start_year = _today.year
    _start_month = 1
else:
    # Other months → next month, last year
    _start_year = _today.year - 1
    _start_month = _today.month + 1

_start_dt = date(_start_year, _start_month, 1)
START_DATE = _start_dt.strftime("%d-%b-%Y")

# New Referrals column fragment (from your HTML)
NEW_REFERRALS_FRAGMENT = "120iT2"

# ===================== DRIVER =====================

def try_login(max_retries=3):
    """
    Attempt to log in up to max_retries times.
    """
    for attempt in range(1, max_retries + 1):
        try:
            login()
            print(f"Login successful on attempt {attempt}")
            return
        except Exception as e:
            print(f"Login attempt {attempt} failed: {e}")
            if attempt < max_retries:
                time.sleep(2)
            else:
                raise

def build_driver():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    opts = Options()
    # Ignore SSL interstitials
    opts.set_capability("acceptInsecureCerts", True)
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--ignore-ssl-errors=yes")

    # Keep Chrome quiet
    opts.add_argument("--log-level=3")
    opts.add_argument("--disable-logging")
    opts.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    opts.add_argument("--disable-features=OptimizationHints,OptimizationGuideModelDownloading,OptimizationGuideOnDeviceModel,SegmentationPlatform")

    # Window
    opts.add_argument("--start-maximized")
    # opts.add_argument("--headless=new")  # optional

    # Auto-download without prompts
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }

    opts.add_experimental_option("prefs", prefs)

    service = Service(log_output=os.devnull)
    driver = webdriver.Chrome(service=service, options=opts)

    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": str(DOWNLOAD_DIR)
        })
    except Exception:
        pass

    return driver

driver = None

def bypass_ssl_interstitial(max_wait=6):
    """Bypass Chrome's SSL warning (Advanced→Proceed or 'thisisunsafe')."""
    end = time.time() + max_wait
    while time.time() < end:
        try:
            details = driver.find_elements(By.ID, "details-button")
            if details:
                details[0].click()
                time.sleep(0.3)
                proceed = driver.find_elements(By.ID, "proceed-link")
                if proceed:
                    proceed[0].click()
                    return
            advanced = driver.find_elements(By.CSS_SELECTOR, "button#more-info, button[aria-label*='Advanced']")
            if advanced:
                advanced[0].click()
                time.sleep(0.3)
                proceed2 = driver.find_elements(By.CSS_SELECTOR, "#proceed-link, a.proceed-link")
                if proceed2:
                    proceed2[0].click()
                    return
        except Exception:
            pass
        time.sleep(0.25)
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys("thisisunsafe")
        time.sleep(0.5)
    except Exception:
        pass

# ===================== SMALL HELPERS =====================
def wait_clickable(by, selector, timeout=100):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, selector)))

def wait_present(by, selector, timeout=100):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, selector)))

def safe_click(el):
    """Scroll into view, try native click, fallback to JS click."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.05)
        el.click()
    except (ElementClickInterceptedException, MoveTargetOutOfBoundsException, WebDriverException):
        driver.execute_script("arguments[0].click();", el)

# ===================== LOGIN & NAV =====================
def login():
    driver.get(LOGIN_URL)
    bypass_ssl_interstitial()
    driver.switch_to.default_content()

    # --- Step 1: Organization ID page ---
    try:
        org_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#orgUserID, input[placeholder='Organization ID']"))
        )

        try:
            driver.execute_script("""
                arguments[0].removeAttribute('readonly');
                arguments[0].readOnly = false;
                arguments[0].disabled = false;
            """, org_box)
        except Exception:
            pass

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", org_box)
        except Exception:
            pass

        try:
            org_box.clear()
        except Exception:
            pass

        try:
            org_box.send_keys(ORG_ID)
        except Exception:
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
                arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
            """, org_box, ORG_ID)

        continue_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnContinue"))
        )
        try:
            continue_btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", continue_btn)

        WebDriverWait(driver, 15).until(EC.any_of(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#username, input[placeholder='User ID']")),
            EC.presence_of_element_located((By.CSS_SELECTOR, "iframe, frame"))
        ))
    except TimeoutException:
        pass

    # --- Step 2: Find username field even if inside an iframe ---
    user_el = None
    pass_el = None

    try:
        user_el = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#username, input[placeholder='User ID']"))
        )
    except TimeoutException:
        if not switch_into_frame_containing("#username, input[placeholder='User ID']", by=By.CSS_SELECTOR, timeout=50, max_depth=10):
            if not switch_into_frame_containing("input[aria-label='User ID']", by=By.CSS_SELECTOR, timeout=50, max_depth=10):
                raise RuntimeError("Could not locate the username field on the login page.")

        user_el = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#username, input[placeholder='User ID'], input[aria-label='User ID']"))
        )

    # --- Step 3: Find password in same frame ---
    try:
        pass_el = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#password, input[placeholder='Password'], input[type='password']"))
        )
    except TimeoutException:
        raise RuntimeError("Found username but not the password field.")

    # --- Step 4: Fill credentials ---
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", user_el)
    except Exception:
        pass
    time.sleep(0.05)

    try:
        user_el.clear()
    except Exception:
        pass
    try:
        user_el.send_keys(USERNAME)
    except Exception:
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
        """, user_el, USERNAME)

    try:
        pass_el.clear()
    except Exception:
        pass
    try:
        pass_el.send_keys(PASSWORD)
    except Exception:
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
        """, pass_el, PASSWORD)

    # --- Step 5: Click Log On ---
    login_btn = None
    selectors = [
        (By.ID, "btnLogOn_TreatAuth"),
        (By.CSS_SELECTOR, "button#btnLogOn_TreatAuth"),
        (By.XPATH, "//button[@type='submit' and .//span[normalize-space()='Log On']]"),
        (By.XPATH, "//button[normalize-space()='Log On' or normalize-space()='Log in' or normalize-space()='Sign in']"),
        (By.CSS_SELECTOR, "form button[type='submit'], form input[type='submit']"),
    ]
    for by, sel in selectors:
        try:
            login_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, sel)))
            break
        except TimeoutException:
            continue

    if not login_btn:
        raise RuntimeError("Login button not found on the credential form.")

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", login_btn)
        time.sleep(0.05)
        login_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", login_btn)

    # --- Step 6: Leave frame and wait for post-login UI ---
    driver.switch_to.default_content()
    try:
        WebDriverWait(driver, 20).until(EC.any_of(
            EC.url_contains("/treat/app"),
            EC.presence_of_element_located((By.CSS_SELECTOR, ".nav-item.dropdown .nav-link.dropdown-toggle")),
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href,'newReports')]"))
        ))
    except TimeoutException:
        print("[DEBUG] After clicking Log On:")
        print("  URL:", driver.current_url)
        print("  Title:", driver.title)
        raise


def handle_active_session_dialog():
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Active Session Detected')]"))
        )
        print("Active session detected. Logging off previous session...")
        btn = wait_clickable(By.XPATH, "//button[contains(text(), 'Log Off Previous Session')]", timeout=10)
        safe_click(btn)
        time.sleep(1.5)
    except TimeoutException:
        pass

def open_synthesis_page():
    dropdowns = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".nav-item.dropdown .nav-link.dropdown-toggle"))
    )
    if len(dropdowns) < 3:
        raise RuntimeError("Reports dropdown not found.")
    safe_click(dropdowns[2])

    synthesis_link = wait_clickable(By.CSS_SELECTOR, "a.dropdown-item[href='/treat/app/reports?page=newReports']")
    before = set(driver.window_handles)
    safe_click(synthesis_link)
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(before))
    for w in driver.window_handles:
        if w not in before:
            driver.switch_to.window(w)
            break

def back_to_synthesis_home(timeout=50):
    """
    Return to the Synthesis report list reliably.
    First try the logo, then fall back to the Synthesis home URL.
    """
    driver.switch_to.default_content()

    home_markers = [
        (By.ID, "report4B"),
        (By.ID, "report30B"),
        (By.ID, "report31B"),
        (By.ID, "report40B"),
    ]

    def _wait_for_home(wait_timeout):
        end = time.time() + wait_timeout
        while time.time() < end:
            driver.switch_to.default_content()
            for by, sel in home_markers:
                try:
                    if driver.find_elements(by, sel):
                        return True
                except Exception:
                    pass
            time.sleep(0.25)
        return False

    try:
        logo = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//img[contains(@src,'Synthesis_logo')]"))
        )
        safe_click(logo)
        if _wait_for_home(8):
            return
    except Exception:
        pass

    driver.get("https://sharedservices.treat.ca/treat/app/reports?page=newReports")
    if _wait_for_home(timeout):
        return

    raise RuntimeError("Could not return to Synthesis home page.")

# ===================== REPORT ACTIONS =====================
def find_report(report_id: str):
    driver.switch_to.default_content()
    el = wait_clickable(By.ID, report_id)
    safe_click(el)

def set_dates():
    s_icon = wait_clickable(By.XPATH, "//input[@id='start']/following-sibling::img[@alt='calendar']")
    safe_click(s_icon)
    s_in = wait_present(By.ID, "start")
    s_in.clear()
    s_in.send_keys(START_DATE)

    e_icon = wait_clickable(By.XPATH, "//input[@id='end']/following-sibling::img[@alt='calendar']")
    safe_click(e_icon)
    e_in = wait_present(By.ID, "end")
    e_in.clear()
    e_in.send_keys(TODAY)

    e_in.send_keys(Keys.ESCAPE)
    try:
        driver.execute_script("document.activeElement && document.activeElement.blur();")
    except Exception:
        pass
    driver.execute_script("document.body.click();")
    time.sleep(0.3)

def set_dates_fiscal():
    """
    Set Start/End to the current fiscal year:
      - Start: Apr 1 of this year if today >= Apr 1; otherwise Apr 1 of last year
      - End: today
    """
    today = datetime.today().date()
    apr1_this_year = date(today.year, 4, 1)
    start_dt = apr1_this_year if today >= apr1_this_year else date(today.year - 1, 4, 1)

    start_str = start_dt.strftime("%d-%b-%Y")
    end_str = today.strftime("%d-%b-%Y")

    s_icon = wait_clickable(By.XPATH, "//input[@id='start']/following-sibling::img[@alt='calendar']")
    safe_click(s_icon)
    s_in = wait_present(By.ID, "start")
    s_in.clear()
    s_in.send_keys(start_str)

    e_icon = wait_clickable(By.XPATH, "//input[@id='end']/following-sibling::img[@alt='calendar']")
    safe_click(e_icon)
    e_in = wait_present(By.ID, "end")
    e_in.clear()
    e_in.send_keys(end_str)

    e_in.send_keys(Keys.ESCAPE)
    try:
        driver.execute_script("document.activeElement && document.activeElement.blur();")
    except Exception:
        pass
    driver.execute_script("document.body.click();")
    time.sleep(0.3)

def choose_all_programs():
    toggle = wait_present(By.ID, "programDiv")
    safe_click(toggle)
    container = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "programCheckBoxDiv")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", container)
    time.sleep(0.2)

    candidates = [
        (By.XPATH, "//*[@id='programCheckBoxDiv']//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'all')]/preceding-sibling::input[@type='checkbox']"),
        (By.CSS_SELECTOR, "#programCheckBoxDiv input[type='checkbox']"),
        (By.XPATH, "//*[@id='q2']/li[1]/input"),
    ]
    for by, sel in candidates:
        try:
            els = container.find_elements(by, sel) if by != By.XPATH or "q2" not in sel else driver.find_elements(by, sel)
        except Exception:
            els = driver.find_elements(by, sel)
        if els:
            safe_click(els[0])
            time.sleep(0.2)
            return
    raise RuntimeError("Could not locate a program 'All' checkbox.")

def click_generate():
    btn = wait_clickable(By.XPATH, "//input[@value='Generate Report']")
    safe_click(btn)
    time.sleep(2)

def download_csv_from_viewer():
    WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "reportViewer")))
    dropdown = wait_present(By.ID, "rvViewer_ctl01_ctl05_ctl00")
    Select(dropdown).select_by_visible_text("CSV (comma delimited)")
    export_btn = wait_clickable(By.ID, "rvViewer_ctl01_ctl05_ctl01")
    safe_click(export_btn)
    driver.switch_to.default_content()
    time.sleep(2)

def click_next_page_in_viewer(wait_secs=2):
    WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "reportViewer")))
    candidates = [
        (By.CSS_SELECTOR, "input[title='Next Page']"),
        (By.CSS_SELECTOR, "input[alt='Next Page']"),
        (By.NAME, "rvViewer$ctl01$ctl01$ctl05$ctl00$ctl00"),
        (By.XPATH, "//input[@type='image' and (contains(@title,'Next') or contains(@alt,'Next'))]"),
    ]
    btn = None
    for by, sel in candidates:
        try:
            btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, sel)))
            break
        except TimeoutException:
            continue
    if not btn:
        driver.switch_to.default_content()
        raise RuntimeError("Next Page button not found.")
    safe_click(btn)
    time.sleep(wait_secs)
    driver.switch_to.default_content()

# --- minimal fragment-based total click ---
def _xpath_for_fragment(fragment: str) -> str:
    return f"//*[contains(@onclick, \"Drillthrough','{fragment}:\")]"

def _switch_into_frame_with_xpath(xpath_expr: str, max_depth: int = 4) -> bool:
    driver.switch_to.default_content()

    def dfs(depth: int) -> bool:
        try:
            if driver.find_elements(By.XPATH, xpath_expr):
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

    return dfs(0)

def click_total_by_fragment(fragment: str = NEW_REFERRALS_FRAGMENT, wait_secs: int = 2, timeout: int = 20):
    """Click the bottom-most drillthrough node for a given column fragment (the Total)."""
    xpath = _xpath_for_fragment(fragment)
    end = time.time() + timeout
    while time.time() < end:
        if _switch_into_frame_with_xpath(xpath):
            break
        time.sleep(0.5)
    else:
        raise RuntimeError(f"No nodes found for fragment {fragment!r} (page not loaded or fragment different).")

    nodes = driver.find_elements(By.XPATH, xpath)
    if not nodes:
        driver.switch_to.default_content()
        raise RuntimeError(f"No nodes found for fragment {fragment!r} after switching to content frame.")

    candidates = []
    for el in nodes:
        rect = driver.execute_script("return arguments[0].getBoundingClientRect();", el)
        top, left = float(rect["top"]), float(rect["left"])
        txt = (el.text or "").strip().replace(",", "")
        num = None
        if any(ch.isdigit() for ch in txt):
            try:
                num = float(txt)
            except Exception:
                num = None
        candidates.append({"el": el, "top": top, "left": left, "num": num})

    max_top = max(c["top"] for c in candidates)
    same_row = [c for c in candidates if (max_top - c["top"]) < 3.0]
    numeric_row = [c for c in same_row if c["num"] is not None]
    target = max(numeric_row, key=lambda c: c["num"]) if numeric_row else sorted(same_row, key=lambda c: c["left"])[0]

    safe_click(target["el"])
    try:
        WebDriverWait(driver, 10).until(EC.staleness_of(target["el"]))
    except TimeoutException:
        pass
    time.sleep(wait_secs)
    driver.switch_to.default_content()

def _switch_into_frame_with_xpath_visible(xpath_expr: str, timeout: int = 20, max_depth: int = 8) -> bool:
    """DFS through frames until an element matching xpath is visible in that frame."""
    end = time.time() + timeout
    while time.time() < end:
        driver.switch_to.default_content()

        def dfs(depth: int) -> bool:
            try:
                els = driver.find_elements(By.XPATH, xpath_expr)
                for el in els:
                    if el.is_displayed():
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
        time.sleep(0.25)
    return False

def _snapshot(dirpath: Path, suffix: str | None = None) -> dict[str, tuple[float, int]]:
    """
    Snapshot files as:
        {filename: (mtime, size)}
    Optionally filter by suffix.
    Ignores in-progress temp files.
    """
    snap = {}
    for f in os.listdir(dirpath):
        if f.endswith(".crdownload") or f.endswith(".tmp"):
            continue
        if suffix and not f.endswith(suffix):
            continue
        p = dirpath / f
        if p.is_file():
            try:
                stat = p.stat()
                snap[f] = (stat.st_mtime, stat.st_size)
            except FileNotFoundError:
                pass
    return snap

def _wait_for_new_file(dirpath: Path, before: dict[str, tuple[float, int]], suffix=".csv", timeout=60) -> Path:
    """
    Wait until either:
      1. a new file appears, or
      2. an existing file is updated (mtime/size changes)
    Return the detected file Path.
    """
    end = time.time() + timeout
    while time.time() < end:
        if any(name.endswith(".crdownload") or name.endswith(".tmp") for name in os.listdir(dirpath)):
            time.sleep(0.25)
            continue

        current = _snapshot(dirpath, suffix=suffix)

        changed = []
        for fname, meta in current.items():
            if fname not in before:
                changed.append(fname)
            elif before[fname] != meta:
                changed.append(fname)

        if changed:
            newest = max(changed, key=lambda fn: (dirpath / fn).stat().st_mtime)
            return dirpath / newest

        time.sleep(0.25)

    raise TimeoutError("No new or updated download detected.")

def _safe_rename(src: Path, dst: Path, add_timestamp_if_exists: bool = False) -> Path:
    src = Path(src)
    dst = Path(dst)

    try:
        if src.resolve(strict=True) == dst.resolve(strict=False):
            if dst.exists():
                return dst
    except FileNotFoundError:
        pass

    dst.parent.mkdir(parents=True, exist_ok=True)

    if dst.exists():
        if add_timestamp_if_exists:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dst = dst.with_name(f"{dst.stem}_{stamp}{dst.suffix}")
        else:
            dst.unlink()

    src.replace(dst)
    return dst

def click_generate_external(timeout: int = 20):
    """
    Clicks the <input type='button' value='Generate Report'> that calls runExternalReport().
    """
    xpath = "//input[@type='button' and @value='Generate Report']"
    btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.1)
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)
    time.sleep(2)

def click_generate_external_in_buttons_container(timeout: int = 20):
    """
    Finds <div class='buttonsContainer'> ... <input value='Generate Report'> ... and triggers it.
    If native click fails, falls back to JS click and finally to calling runExternalReport(...) directly.
    """
    btn_xpath = "//div[contains(@class,'buttonsContainer')]//input[@type='button' and @value='Generate Report']"

    driver.switch_to.default_content()

    if not _switch_into_frame_with_xpath_visible(btn_xpath, timeout=timeout):
        raise RuntimeError("Could not find a visible 'Generate Report' button inside .buttonsContainer.")

    btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.05)
        btn.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            driver.execute_script("if (typeof runExternalReport === 'function') { runExternalReport(310, 'reportViewerContainer', true); }")
    time.sleep(2)

    driver.switch_to.default_content()
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "reportViewerContainer"))
        )
    except TimeoutException:
        pass

# ========= Generic Frame Finder =========
def switch_into_frame_containing(selector, by=By.CSS_SELECTOR, timeout=50, max_depth=10):
    """
    DFS into nested frames until an element matching (by, selector) exists.
    Leaves driver focused in the frame that contains it. Returns True/False.
    """
    end = time.time() + timeout
    while time.time() < end:
        driver.switch_to.default_content()

        def dfs(depth):
            try:
                if driver.find_elements(by, selector):
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
        time.sleep(0.25)
    return False

# ========= Exporters for common viewers =========
def try_export_ssrs_csv(timeout=15):
    """
    SSRS viewer export: choose 'CSV (comma delimited)' and click Export.
    """
    selects = driver.find_elements(By.TAG_NAME, "select")
    target_sel = None
    for sel in selects:
        try:
            opts = [o.text.strip().lower() for o in sel.find_elements(By.TAG_NAME, "option")]
            if any("csv" in t for t in opts):
                target_sel = sel
                break
        except Exception:
            continue

    if not target_sel:
        return False

    Select(target_sel).select_by_visible_text("CSV (comma delimited)")
    try:
        export_btn = driver.find_element(
            By.XPATH,
            ".//input[@type='submit' or @type='image' or @type='button']"
            "[contains(@id,'ctl05_ctl01') or contains(@title,'Export') or contains(@alt,'Export')]"
        )
    except Exception:
        buttons = [b for b in driver.find_elements(By.XPATH, "//input|//button") if b.is_displayed()]
        export_btn = buttons[0] if buttons else None
    if not export_btn:
        return False

    safe_click(export_btn)
    return True

def try_export_webfocus_csv(timeout=15):
    """
    WebFOCUS viewer/file grid export: click Export then choose CSV.
    """
    candidates = [
        (By.XPATH, "//*[contains(@aria-label,'Export') or contains(@title,'Export')][self::button or self::div or self::span or self::a]"),
        (By.CSS_SELECTOR, "[data-ibx-command='cmdExport']"),
        (By.XPATH, "//div[contains(@class,'ibx-menu-item')][.//div[contains(@class,'ibx-label-text')][contains(translate(.,'CSV','csv'),'csv')]]"),
    ]

    clicked_export = False
    for by, sel in candidates:
        try:
            el = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((by, sel)))
            safe_click(el)
            time.sleep(0.3)
            clicked_export = True
            break
        except Exception:
            continue

    csv_selectors = [
        (By.XPATH, "//*[contains(translate(normalize-space(.),'CSV','csv'),'csv')][self::div or self::span or self::a or self::button]"),
        (By.CSS_SELECTOR, "div.ibx-menu-item, a, button, span"),
    ]
    time.sleep(0.2)
    for by, sel in csv_selectors:
        items = [e for e in driver.find_elements(by, sel) if e.is_displayed()]
        for it in items:
            try:
                label = (it.text or it.get_attribute("aria-label") or it.get_attribute("title") or "").strip().lower()
                if "csv" in label:
                    safe_click(it)
                    return True
            except Exception:
                continue

    try:
        a_csv = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(translate(@href,'CSV','csv'),'.csv')]"))
        )
        safe_click(a_csv)
        return True
    except Exception:
        pass

    return clicked_export

def export_csv_any_viewer():
    """
    Find a plausible viewer area, then attempt SSRS export, else WebFOCUS export.
    """
    driver.switch_to.default_content()

    viewer_hints = [
        (By.ID, "reportViewer"),
        (By.ID, "ReportViewer"),
        (By.ID, "reportViewerContainer"),
        (By.CSS_SELECTOR, "iframe[id*='report'], iframe[id*='viewer'], frame[id*='report']"),
        (By.CSS_SELECTOR, "div[id*='viewer']"),
    ]

    found = False
    for by, sel in viewer_hints:
        if switch_into_frame_containing(sel, by=by, timeout=10):
            found = True
            break

    try:
        if try_export_ssrs_csv():
            driver.switch_to.default_content()
            return True
    except Exception:
        pass

    try:
        if try_export_webfocus_csv():
            driver.switch_to.default_content()
            return True
    except Exception:
        pass

    driver.switch_to.default_content()
    return False

def wait_for_download(timeout=240):
    end = time.time() + timeout
    while time.time() < end:
        if any(name.endswith(".crdownload") for name in os.listdir(DOWNLOAD_DIR)):
            time.sleep(0.4)
            continue
        files = [os.path.join(DOWNLOAD_DIR, f) for f in os.listdir(DOWNLOAD_DIR) if os.path.isfile(os.path.join(DOWNLOAD_DIR, f))]
        if files:
            newest = max(files, key=os.path.getmtime)
            return newest
        time.sleep(0.4)
    raise TimeoutError("Download did not finish in time.")

# ========= Retryable tile double-clicker (optional) =========
def try_double_click_tile(tile_label: str, max_retries=3, timeout=25):
    """
    Attempt to double-click a file-grid tile with given label text, retrying on failure.
    """
    for attempt in range(1, max_retries + 1):
        try:
            double_click_tile(tile_label, timeout=timeout)
            print(f"Double-clicked '{tile_label}' on attempt {attempt}")
            return
        except Exception as e:
            print(f"Double-click attempt {attempt} for '{tile_label}' failed: {e}")
            if attempt < max_retries:
                time.sleep(2)
            else:
                raise

def double_click_tile(label_text: str, timeout=25, max_depth=8):
    """
    Locate and double-click a WebFOCUS-style file grid tile by its label text.
    """
    label_xpath = f"//div[contains(@class,'ibx-label-text')][normalize-space()='{label_text}']"

    def switch_into_frame_with_visible(xpath_expr: str) -> bool:
        end = time.time() + timeout
        while time.time() < end:
            driver.switch_to.default_content()

            def dfs(depth: int) -> bool:
                try:
                    for el in driver.find_elements(By.XPATH, xpath_expr):
                        if el.is_displayed():
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
            time.sleep(0.2)
        return False

    if not switch_into_frame_with_visible(label_xpath):
        driver.switch_to.default_content()
        raise RuntimeError(f"Couldn't find visible '{label_text}' in any frame.")

    label = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, label_xpath)))
    try:
        clickable = label.find_element(By.XPATH, "./ancestor::div[contains(@class,'image-text')][1]")
    except Exception:
        clickable = label.find_element(By.XPATH, "./ancestor::div[contains(@class,'file-item')][1]")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", clickable)
    time.sleep(0.05)

    try:
        ActionChains(driver).move_to_element(clickable).double_click(clickable).perform()
    except Exception:
        driver.execute_script("""
            const el = arguments[0];
            el.scrollIntoView({block:'center'});
            const evt = new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window});
            el.dispatchEvent(evt);
        """, clickable)

    driver.switch_to.default_content()
    time.sleep(1.0)

# ========= Generic context-menu "Run" on a tile =========
def run_tile_via_context_menu(tile_label: str, timeout: int = 25, max_depth: int = 8):
    """
    Right-click the given tile label and click the 'Run' menu item
    identified by data-ibx-command='cmdRun' / action='run'.
    """
    label_xpath = f"//div[contains(@class,'ibx-label-text')][normalize-space()='{tile_label}']"
    run_xpath = ("//div[@data-ibx-type='ibxMenuItem' "
                 "and (translate(@data-ibx-command,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='cmdrun' "
                 "or translate(@action,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='run')]"
                 "[.//div[contains(@class,'ibx-label-text')][normalize-space()='Run']]")

    end = time.time() + timeout
    found = False
    while time.time() < end and not found:
        driver.switch_to.default_content()

        def dfs(depth: int) -> bool:
            try:
                for el in driver.find_elements(By.XPATH, label_xpath):
                    if el.is_displayed():
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
            found = True
        else:
            time.sleep(0.2)

    if not found:
        driver.switch_to.default_content()
        raise RuntimeError(f"Couldn't find visible '{tile_label}' in any frame.")

    label = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, label_xpath)))
    try:
        tile = label.find_element(By.XPATH, "./ancestor::div[contains(@class,'image-text')][1]")
    except Exception:
        tile = label.find_element(By.XPATH, "./ancestor::div[contains(@class,'file-item')][1]")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", tile)
    time.sleep(0.05)
    try:
        tile.click()
        time.sleep(0.1)
    except Exception:
        driver.execute_script("arguments[0].click();", tile)
        time.sleep(0.1)

    try:
        ActionChains(driver).context_click(tile).perform()
    except Exception:
        try:
            tile.send_keys(Keys.SHIFT, Keys.F10)
        except Exception:
            driver.execute_script("""
                const el = arguments[0];
                const evt = new MouseEvent('contextmenu', {bubbles:true, cancelable:true, view:window, buttons:2});
                el.dispatchEvent(evt);
            """, tile)
    time.sleep(0.2)

    try:
        run_item = WebDriverWait(driver, 6).until(
            EC.visibility_of_element_located((By.XPATH, run_xpath))
        )
    except TimeoutException:
        try:
            run_item = WebDriverWait(driver, 4).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR,
                    "div.ibx-menu-item[data-ibx-command='cmdRun'], div.ibx-menu-item[action='run']"))
            )
        except TimeoutException:
            raise RuntimeError("Context menu opened, but 'Run' was not found.")

    aria_disabled = (run_item.get_attribute("aria-disabled") or "false").lower()
    if aria_disabled == "true":
        raise RuntimeError("'Run' menu item is disabled.")

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", run_item)
        time.sleep(0.05)
        run_item.click()
    except Exception:
        driver.execute_script("arguments[0].click();", run_item)

    driver.switch_to.default_content()
    time.sleep(1.0)

def _dispatch_change(driver, el):
    driver.execute_script("""
        const el = arguments[0];
        el.dispatchEvent(new Event('input', {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
    """, el)

def _set_autop_date_inputs(start_str: str, end_str: str, timeout: int = 20) -> None:
    """
    Find the WebFOCUS autop date inputs by their placeholders, remove 'readonly',
    set values, and fire input/change events. Searches across frames.
    """
    if not switch_into_frame_containing("input[placeholder^='Start Date']", by=By.CSS_SELECTOR, timeout=timeout):
        raise RuntimeError("Start Date input not found.")
    s_in = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder^='Start Date']"))
    )
    driver.execute_script("arguments[0].removeAttribute('readonly');", s_in)
    s_in.clear()
    s_in.send_keys(start_str)
    _dispatch_change(driver, s_in)

    try:
        e_in = WebDriverWait(driver, 3).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder^='End Date']"))
        )
    except TimeoutException:
        if not switch_into_frame_containing("input[placeholder^='End Date']", by=By.CSS_SELECTOR, timeout=timeout):
            raise RuntimeError("End Date input not found.")
        e_in = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder^='End Date']"))
        )

    driver.execute_script("arguments[0].removeAttribute('readonly');", e_in)
    e_in.clear()
    e_in.send_keys(end_str)
    _dispatch_change(driver, e_in)

    driver.switch_to.default_content()

def _select_export_option(label_text: str = "Report Export (XLSX)", timeout: int = 50) -> None:
    """
    Open the export dropdown (the one whose visible label contains 'Report Export')
    and choose the given label.
    """
    btn_xpath = "//a[span[contains(@class,'autop-control') and contains(normalize-space(.),'Report Export')]]"
    if not _switch_into_frame_with_xpath_visible(btn_xpath, timeout=timeout):
        raise RuntimeError("Export dropdown button not visible.")
    btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
    safe_click(btn)
    time.sleep(0.2)

    opt_xpath = f"//*[contains(@class,'ui-selectmenu') or @role='option' or self::a or self::div or self::li][normalize-space()='{label_text}']"
    opts = driver.find_elements(By.XPATH, opt_xpath)
    if not any(o.is_displayed() for o in opts):
        driver.switch_to.default_content()
        opts = [o for o in driver.find_elements(By.XPATH, opt_xpath) if o.is_displayed()]
        if opts:
            safe_click(opts[0])
        else:
            fallback = [o for o in driver.find_elements(By.XPATH, f"//*[normalize-space()='{label_text}']") if o.is_displayed()]
            if not fallback:
                raise RuntimeError(f"Export option '{label_text}' not found.")
            safe_click(fallback[0])
    else:
        safe_click([o for o in opts if o.is_displayed()][0])

    driver.switch_to.default_content()

def _click_autop_run(timeout: int = 20) -> None:
    """
    Click the 'Run with filter values' button.
    """
    run_xpath = "//a[contains(@class,'autop-btn-submit') and contains(@title,'Run')]"
    if not _switch_into_frame_with_xpath_visible(run_xpath, timeout=timeout):
        raise RuntimeError("Run button not visible.")
    run_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, run_xpath)))
    safe_click(run_btn)
    driver.switch_to.default_content()

def run_census_report_webfocus_yearly():
    """
    WebFOCUS → 'Census Report' with calendar-year dates:
    Start = Jan 1 of current year, End = today. Exports XLSX.
    """
    find_report("report40B")
    click_generate_external()
    click_generate_external_in_buttons_container(timeout=360)

    run_tile_via_context_menu("Census Report", timeout=360)

    today = datetime.today().date()
    start_dt = date(today.year, 1, 1)
    start_str = start_dt.strftime("%d-%b-%Y")
    end_str = today.strftime("%d-%b-%Y")

    _set_autop_date_inputs(start_str, end_str)
    _select_export_option("Report Export (XLSX)")

    before = _snapshot(DOWNLOAD_DIR, suffix=".xlsx")
    _click_autop_run()
    xlsx_path = _wait_for_new_file(DOWNLOAD_DIR, before, suffix=".xlsx", timeout=360)
    _safe_rename(xlsx_path, DOWNLOAD_DIR / "rpt_Census_WebFocus.xlsx", add_timestamp_if_exists=True)
    print(f"[Census Report] Downloaded: {DOWNLOAD_DIR / 'rpt_Census_WebFocus.xlsx'}")

    back_to_synthesis_home()
    time.sleep(1)

# ===================== MAIN FLOW =====================
def run():
    global driver
    driver = build_driver()
    try_login(max_retries=3)
    open_synthesis_page()

    # --- Report 4B ---
    print("Running report 4B...")
    find_report("report4B")
    set_dates()
    click_generate()
    download_csv_from_viewer()
    back_to_synthesis_home()
    time.sleep(1)

    ## --- Report 8B (Census) ---
    #print("Running report 8B (Census)...")
    #find_report("report8B")
    #set_dates()
    #choose_all_programs()
    #click_generate()
    #download_csv_from_viewer()
    #back_to_synthesis_home()
    #time.sleep(1)
#
    # --- Report 30B ---
    print("Running report 30B...")
    find_report("report30B")
    set_dates()
    choose_all_programs()
    click_generate()
    download_csv_from_viewer()
    back_to_synthesis_home()
    time.sleep(1)

    # --- Report 41B ---
    print("Running report 41B...")
    find_report("report41B")
    set_dates()
    click_generate()
    download_csv_from_viewer()
    back_to_synthesis_home()
    time.sleep(1)

    # --- Report 17B (MIS - Statistics) ---
    print("Running report 17B (MIS - Statistics)...")
    find_report("report17B")
    set_dates_fiscal()
    choose_all_programs()
    click_generate()
    download_csv_from_viewer()
    back_to_synthesis_home()
    time.sleep(1)

        # --- Report 31B New Referrals ---
    print("Running report 31B (Referral Reports New Referrals)...")
    find_report("report31B")
    set_dates()
    choose_all_programs()
    click_generate()
#
    click_next_page_in_viewer()
    click_total_by_fragment(fragment=NEW_REFERRALS_FRAGMENT, wait_secs=3, timeout=250)
#
    before = _snapshot(DOWNLOAD_DIR, suffix=".csv")
    download_csv_from_viewer()
    first_path = _wait_for_new_file(DOWNLOAD_DIR, before, suffix=".csv", timeout=240)
#
    first_final = DOWNLOAD_DIR / "rpt_Referrals_DT.csv"
    _safe_rename(first_path, first_final, add_timestamp_if_exists=False)
#
    # Temporarily move first file away so Chrome cannot overwrite it
    temp_hold = DOWNLOAD_DIR / "__temp_rpt_Referrals_DT.csv"
    if temp_hold.exists():
        temp_hold.unlink()
    first_final.replace(temp_hold)
#
    back_to_synthesis_home(timeout=20)
    time.sleep(2)
#
    # --- 31B — Is Waitlisted total ---
    print("Running report 31B (Census) — Is Waitlisted...")
    find_report("report31B")
    set_dates()
    choose_all_programs()
    click_generate()
    click_next_page_in_viewer()
#
    WAITLISTED_FRAGMENT = "132iT2"
    click_total_by_fragment(fragment=WAITLISTED_FRAGMENT, wait_secs=2, timeout=50)
#
    before = _snapshot(DOWNLOAD_DIR, suffix=".csv")
    download_csv_from_viewer()
    second_path = _wait_for_new_file(DOWNLOAD_DIR, before, suffix=".csv", timeout=240)
#
    second_final = DOWNLOAD_DIR / "rpt_Referrals_Waitlisted.csv"
    _safe_rename(second_path, second_final, add_timestamp_if_exists=False)
#
    # Move first file back after second download is safely renamed
    temp_hold.replace(first_final)
#
    back_to_synthesis_home(timeout=20)
    time.sleep(2)

    # --- WebFOCUS ---
    print("Running WebFOCUS tiles in sequence...")
    webfocus_tiles = ["Assessments", "External Documents", "Demographics_Active & Discharged", "Note Content RAC","IPCC Goal","IPCC Issue"]

    for idx, tile_label in enumerate(webfocus_tiles):
        find_report("report40B")
        click_generate_external()
        click_generate_external_in_buttons_container(timeout=360)

        print(f"Running tile: {tile_label}")
        run_tile_via_context_menu(tile_label, timeout=360)

        if not export_csv_any_viewer():
            print(f"[{tile_label}] Could not find an Export→CSV control automatically.")
        else:
            try:
                downloaded_path = wait_for_download(timeout=240)
                print(f"[{tile_label}] Downloaded: {downloaded_path}")
            except Exception as e:
                print(f"[{tile_label}] No download detected: {e}")

        if idx < len(webfocus_tiles) - 1:
            back_to_synthesis_home()
            time.sleep(1)

if __name__ == "__main__":
    try:
        run()
    finally:
        time.sleep(5)
        driver.quit()