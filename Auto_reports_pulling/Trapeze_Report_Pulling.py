# -*- coding: utf-8 -*-
# Trapeze end-to-end pull — credentials from environment variables
# pip install selenium webdriver-manager

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
    ElementClickInterceptedException,
)
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

from datetime import datetime, timedelta
from pathlib import Path
import time, os, shutil, re

# =========================
# CONFIG
# =========================
URL = "https://trapeze.sprintseniorcare.org/App/core/ui/workspace/main"

DOWNLOAD_DIR = None  # or r"C:\Temp\trapeze_downloads"
HEADLESS = False     # True to run headless
KEEP_OPEN = False    # True to keep window open (detach)

INCOMPLETE_EXTS = {".crdownload", ".partial", ".tmp"}

# =========================
# Credentials from environment
# =========================
def _get_creds_from_env():
    u = os.getenv("TRAPEZE_USERNAME", "").strip()
    p = os.getenv("TRAPEZE_PASSWORD", "")
    if not u or not p:
        raise RuntimeError(
            "Missing credentials. Set TRAPEZE_USERNAME and TRAPEZE_PASSWORD in your environment."
        )
    return u, p

# =========================
# Paths & date helpers
# =========================
def _home() -> Path:
    return Path(os.path.expanduser("~")).resolve()

def _downloads_dir() -> Path:
    return _home() / "Downloads"

def _detect_onedrive_root() -> Path | None:
    for key in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.getenv(key)
        if p and Path(p).exists():
            return Path(p)
    for c in [
        _home() / "OneDrive - Reconnect Community Health Services (1)",
        _home() / "OneDrive - Reconnect Community Health Services",
        _home() / "OneDrive",
    ]:
        if c.exists():
            return c
    return None

def _target_dir() -> Path:
    od = _detect_onedrive_root()
    if not od:
        raise RuntimeError("Could not locate OneDrive folder. Set OneDrive env var or create it.")
    t = od / r"data\Clean Data\Trapeze"
    t.mkdir(parents=True, exist_ok=True)
    return t

def previous_full_week_monday_sunday(today=None):
    if today is None:
        today = datetime.today()
    this_monday = today - timedelta(days=today.weekday())
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)
    return last_monday.strftime("%m/%d/%Y"), last_sunday.strftime("%m/%d/%Y")

def _fmt_date_for_name(s: str) -> str:
    s = s.strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%m-%d-%Y")
        except ValueError:
            pass
    return re.sub(r"[^\d\-]", "-", s).strip("-")

def _latest_finished_file(d: Path) -> Path | None:
    files = [
        p for p in d.glob("*")
        if p.is_file()
        and p.suffix.lower() not in INCOMPLETE_EXTS
        and not p.name.startswith(("~", "."))
    ]
    if not files:
        return None
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0]

def _wait_for_download_complete(downloads: Path, timeout: int = 180, quiet_stabilize: float = 1.0) -> Path:
    end = time.time() + timeout
    last_seen: Path | None = None
    last_size = -1
    stable_since = None
    while time.time() < end:
        cand = _latest_finished_file(downloads)
        if cand:
            size = cand.stat().st_size
            if last_seen is None or cand != last_seen:
                last_seen = cand
                last_size = size
                stable_since = time.time()
            else:
                if size == last_size:
                    if stable_since and (time.time() - stable_since) >= quiet_stabilize:
                        return cand
                else:
                    last_size = size
                    stable_since = time.time()
        time.sleep(0.5)
    raise TimeoutException(f"Timed out waiting for a completed download in: {downloads}")

def move_and_rename_report(start_date: str, end_date: str) -> Path:
    downloads = _downloads_dir()
    target = _target_dir()

    src = _wait_for_download_complete(downloads, timeout=180, quiet_stabilize=1.0)

    sd = _fmt_date_for_name(start_date)
    ed = _fmt_date_for_name(end_date)
    dest = (target / f"Trapeze_{sd}_{ed}").with_suffix(src.suffix)

    try:
        if dest.exists():
            dest.unlink()
        shutil.move(str(src), str(dest))
    except PermissionError as e:
        raise PermissionError(f"Cannot replace '{dest}'. Is it open in Excel?") from e

    print(f"Moved & replaced:\n  {src}\n→ {dest}")
    return dest

# =========================
# Selenium helpers
# =========================
def smart_click(driver, el):
    try:
        el.click()
        return True
    except Exception:
        pass
    try:
        ActionChains(driver).move_to_element(el).pause(0.15).click(el).perform()
        return True
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False

def switch_default(driver):
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

def all_iframes_recursive(driver):
    """Yield iframe elements (DFS). None => top document."""
    stack = [None]
    while stack:
        fr = stack.pop()
        switch_default(driver)
        if fr is not None:
            try:
                driver.switch_to.frame(fr)
            except Exception:
                continue
        yield fr
        try:
            children = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            children = []
        for ch in children:
            stack.append(ch)

def find_any_frame(driver, by, value, timeout=15):
    end = time.time() + timeout
    while time.time() < end:
        for _fr in all_iframes_recursive(driver):
            try:
                el = WebDriverWait(driver, 0.9).until(EC.presence_of_element_located((by, value)))
                return el, _fr
            except Exception:
                continue
    raise TimeoutException(f"Element not found: {by}={value}")

def click_any_frame(driver, by, value, timeout=15):
    end = time.time() + timeout
    last_err = None
    while time.time() < end:
        for _fr in all_iframes_recursive(driver):
            try:
                el = WebDriverWait(driver, 1.0).until(EC.element_to_be_clickable((by, value)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                if smart_click(driver, el):
                    return True
            except Exception as e:
                last_err = e
                continue
        time.sleep(0.2)
    raise TimeoutException(f"Could not click: {by}={value}. Last error: {last_err}")

def is_report_menu_open(driver):
    report_xps = [
        "//*[normalize-space()='Trip Export Report for CIMS']",
    ]
    for _fr in all_iframes_recursive(driver):
        for xp in report_xps:
            try:
                els = driver.find_elements(By.XPATH, xp)
                for el in els:
                    try:
                        if el.is_displayed():
                            return True
                    except Exception:
                        continue
            except Exception:
                continue
    return False

def click_recent_footer_tab(driver, timeout=20):
    recent_xps = [
        "//div[contains(@class,'WorkspaceFootItem')][normalize-space()='Recent']",
        "//*[contains(@class,'WorkspaceFootItem') and normalize-space()='Recent']",
        "//*[normalize-space()='Recent' and contains(@class,'WorkspaceFootItem')]",
    ]

    end = time.time() + timeout
    last_err = None

    while time.time() < end:
        if is_report_menu_open(driver):
            return True

        for _fr in all_iframes_recursive(driver):
            for xp in recent_xps:
                try:
                    els = driver.find_elements(By.XPATH, xp)
                except Exception as e:
                    last_err = e
                    continue

                for el in els:
                    try:
                        if not el.is_displayed():
                            continue

                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        time.sleep(0.15)

                        clicked = False

                        try:
                            el.click()
                            clicked = True
                        except Exception:
                            pass

                        if not clicked:
                            try:
                                ActionChains(driver).move_to_element(el).pause(0.2).click(el).perform()
                                clicked = True
                            except Exception:
                                pass

                        if not clicked:
                            try:
                                driver.execute_script("arguments[0].click();", el)
                                clicked = True
                            except Exception:
                                pass

                        if not clicked:
                            try:
                                ActionChains(driver).move_to_element(el).pause(0.2).double_click(el).perform()
                                clicked = True
                            except Exception:
                                pass

                        if clicked:
                            time.sleep(0.6)
                            if is_report_menu_open(driver):
                                return True

                    except Exception as e:
                        last_err = e
                        continue

        time.sleep(0.4)

    raise TimeoutException(f"'Recent' was found but did not expand properly. Last error: {last_err}")

def find_date_inputs(driver, timeout=20):
    """
    Return (start_input, end_input) for <input type='text' size='11'> fields.
    Picks the first two visible ones, across frames.
    """
    end = time.time() + timeout
    xps = [
        "//input[@type='text' and @size='11' and @autocomplete='off']",
        "//input[@type='text' and @size='11']",
    ]
    while time.time() < end:
        for _fr in all_iframes_recursive(driver):
            visibles = []
            for xp in xps:
                try:
                    els = driver.find_elements(By.XPATH, xp)
                except Exception:
                    els = []
                for e in els:
                    try:
                        if e.is_displayed():
                            visibles.append(e)
                    except StaleElementReferenceException:
                        continue
            uniq = []
            seen = set()
            for e in visibles:
                key = getattr(e, "id", None)
                if key and key in seen:
                    continue
                if key:
                    seen.add(key)
                uniq.append(e)
            if len(uniq) >= 2:
                return uniq[0], uniq[1]
        time.sleep(0.2)
    raise TimeoutException("Could not find two visible date inputs (type=text, size=11).")

def set_input_value(driver, el, value):
    try:
        el.click()
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.BACKSPACE)
        el.send_keys(value)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input',{bubbles:true}))", el)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change',{bubbles:true}))", el)
        driver.execute_script("arguments[0].blur && arguments[0].blur();", el)
        return True
    except Exception:
        try:
            driver.execute_script("""
                const el = arguments[0], val = arguments[1];
                el.value = val;
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
                el.dispatchEvent(new Event('blur',{bubbles:true}));
            """, el, value)
            return True
        except Exception:
            return False

def check_service_provider(driver, timeout=10):
    """
    Check the visible service provider field.
    If its value is not exactly '9', print a warning only.
    Do not change anything.
    """
    end = time.time() + timeout

    xps = [
        "//div[contains(@class,'in') and @contenteditable='false']",
    ]

    while time.time() < end:
        for _fr in all_iframes_recursive(driver):
            for xp in xps:
                try:
                    els = driver.find_elements(By.XPATH, xp)
                except Exception:
                    continue

                for el in els:
                    try:
                        if not el.is_displayed():
                            continue

                        text = (el.text or "").strip()
                        if not text:
                            continue

                        # Only react when it looks like a provider selection value
                        # such as 9 or 4,5
                        if re.fullmatch(r"[\d,\s]+", text):
                            if text != "9":
                                print("WARNING: Service Provider is not Reconnect, it might cause error in the Pipeline, please manually change it back.")
                            return
                    except Exception:
                        continue

        time.sleep(0.2)

def wait_new_window(driver, prev_handles, timeout=8):
    end = time.time() + timeout
    while time.time() < end:
        if len(driver.window_handles) > len(prev_handles):
            return True
        time.sleep(0.2)
    return False

def click_generated_report_link(driver, timeout=90):
    link_xps = [
        "//a[normalize-space(.)='Please click here to obtain generated report.']",
        "//a[contains(normalize-space(.), 'Please click here') and contains(normalize-space(.), 'generated report')]",
    ]
    end = time.time() + timeout
    while time.time() < end:
        for _fr in all_iframes_recursive(driver):
            for xp in link_xps:
                try:
                    el = WebDriverWait(driver, 0.9).until(EC.presence_of_element_located((By.XPATH, xp)))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    if smart_click(driver, el):
                        return True
                except Exception:
                    continue
        time.sleep(0.4)
    raise TimeoutException("Generated report link did not appear in time.")

# =========================
# Driver
# =========================
def _build_driver(download_dir=None, headless=False, keep_open=False):
    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1600,1000")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    if keep_open:
        opts.add_experimental_option("detach", True)
    if download_dir:
        Path(download_dir).mkdir(parents=True, exist_ok=True)
        opts.add_experimental_option("prefs", {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    driver.set_page_load_timeout(120)
    return driver

# =========================
# Core run
# =========================
def _run_once(url, username, password, start_date, end_date, download_dir=None, headless=False, keep_open=False):
    driver = _build_driver(download_dir, headless, keep_open)
    try:
        wait = WebDriverWait(driver, 30)
        driver.get(url)

        # --- Login ---
        wait.until(EC.element_to_be_clickable((By.ID, "loginName"))).send_keys(username)
        wait.until(EC.element_to_be_clickable((By.ID, "loginPassword"))).send_keys(password)
        smart_click(driver, wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.BtnLogin[name='login']"))))
        time.sleep(1.5)

        # --- Click "Recent" footer tab ---
        click_recent_footer_tab(driver, timeout=20)

        # --- Open "Trip Export Report for CIMS" (tag-agnostic) ---
        prev_handles = set(driver.window_handles)
        click_any_frame(driver, By.XPATH, "//*[normalize-space()='Trip Export Report for CIMS']", timeout=25)

        if wait_new_window(driver, prev_handles, timeout=6):
            new_handle = list(set(driver.window_handles) - prev_handles)[0]
            driver.switch_to.window(new_handle)

        # --- Fill dates ---
        for attempt in range(3):
            try:
                start_input, end_input = find_date_inputs(driver, timeout=20)

                if not set_input_value(driver, start_input, start_date):
                    start_input, _ = find_date_inputs(driver, timeout=10)
                    if not set_input_value(driver, start_input, start_date):
                        raise TimeoutException("Failed to set Start Date")

                time.sleep(0.3)
                _, end_input = find_date_inputs(driver, timeout=10)
                if not set_input_value(driver, end_input, end_date):
                    _, end_input = find_date_inputs(driver, timeout=10)
                    if not set_input_value(driver, end_input, end_date):
                        raise TimeoutException("Failed to set End Date")

                break
            except StaleElementReferenceException:
                if attempt == 2:
                    raise
                time.sleep(0.25)

        # --- Warn if Service Provider is not 9 ---
        check_service_provider(driver, timeout=10)

        # --- Click "Report" (robust locators) ---
        report_locators = [
            (By.XPATH, "//*[(self::button or self::div or self::span or self::a) "
                       "and contains(@class,'BarItem') and normalize-space()='Report']"),
            (By.XPATH, "//button[contains(@onclick,'3000018')]"),
            (By.XPATH, "//*[normalize-space()='Report']"),
        ]
        clicked = False
        for by, xp in report_locators:
            try:
                click_any_frame(driver, by, xp, timeout=15)
                clicked = True
                break
            except Exception:
                continue
        if not clicked:
            raise TimeoutException("Couldn't click Report button.")

        # --- Click generated report link ---
        click_generated_report_link(driver, timeout=90)

        # --- Move & rename the downloaded file ---
        try:
            move_and_rename_report(start_date, end_date)
        except Exception as e:
            print(f"WARNING: download move/rename failed: {e}")

        msg = f"Downloaded Trapeze report for {start_date} → {end_date}."
        print(msg)
        return msg

    finally:
        if not KEEP_OPEN:
            try:
                driver.quit()
            except Exception:
                pass

# =========================
# Public API
# =========================
def run(
    url: str = URL,
    start_date: str | None = None,
    end_date: str | None = None,
    download_dir: str | None = DOWNLOAD_DIR,
    headless: bool = HEADLESS,
    keep_open: bool = KEEP_OPEN,
):
    username, password = _get_creds_from_env()
    if not start_date or not end_date:
        start_date, end_date = previous_full_week_monday_sunday()
    return _run_once(url, username, password, start_date, end_date, download_dir, headless, keep_open)

# =========================
# CLI
# =========================
def main():
    sd, ed = previous_full_week_monday_sunday()
    run(start_date=sd, end_date=ed)

if __name__ == "__main__":
    main()