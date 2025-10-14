# pip install selenium webdriver-manager

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
import time
import os
from pathlib import Path
import shutil
import re


URL = "https://trapeze.sprintseniorcare.org/App/core/ui/workspace/main"
# --- remove these ---
# USERNAME = "fsijinardo"
# PASSWORD = "ThisSystemMakesMeCrazy2025"

# --- add this near the top (after imports) ---
def _get_trapeze_creds():
    u = os.getenv("TRAPEZE_USERNAME", "").strip()
    p = os.getenv("TRAPEZE_PASSWORD", "")
    if not u or not p:
        raise RuntimeError(
            "Trapeze credentials are missing. Set them in the app (Update Credentials → Save) "
            "or export TRAPEZE_USERNAME / TRAPEZE_PASSWORD."
        )
    return u, p

# ----------------- date helpers -----------------
INCOMPLETE_EXTS = {".crdownload", ".partial", ".tmp"}

def _home() -> Path:
    # More reliable than os.getlogin()
    return Path(os.path.expanduser("~")).resolve()

def _downloads_dir() -> Path:
    return _home() / "Downloads"

def _detect_onedrive_root() -> Path | None:
    # Prefer env vars if available (most reliable)
    for key in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.getenv(key)
        if p and Path(p).exists():
            return Path(p)
    # Fallbacks under the user profile
    candidates = [
        _home() / "OneDrive - Reconnect Community Health Services (1)",
        _home() / "OneDrive - Reconnect Community Health Services",
        _home() / "OneDrive",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None

def _target_dir() -> Path:
    od = _detect_onedrive_root()
    if not od:
        raise RuntimeError("Could not locate OneDrive folder. Set OneDriveCommercial/OneDrive env var or create the folder.")
    t = od / r"data\Clean Data\Trapeze"
    t.mkdir(parents=True, exist_ok=True)
    return t

def _fmt_date_for_name(s: str) -> str:
    s = s.strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%Y/%m/%d"):
        try:
            from datetime import datetime as _dt
            return _dt.strptime(s, fmt).strftime("%m-%d-%Y")
        except ValueError:
            pass
    # last-resort sanitize
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

def _wait_for_download_complete(downloads: Path, timeout: int = 120, quiet_stabilize: float = 1.0) -> Path:
    """
    Wait for the most recent file in Downloads to exist, not be a temp extension,
    and have a stable size for quiet_stabilize seconds. Returns that Path.
    """
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

def _unique_path(dirpath: Path, basename: str, ext: str) -> Path:
    cand = dirpath / f"{basename}{ext}"
    if not cand.exists():
        return cand
    i = 1
    while True:
        cand = dirpath / f"{basename} ({i}){ext}"
        if not cand.exists():
            return cand
        i += 1

def move_and_rename_report(start_date: str, end_date: str) -> Path:
    """
    Find newest finished file in ~/Downloads, move → OneDrive/.../Trapeze,
    rename to: Trapeze_<STARTDATE>_<ENDDATE><ext>.
    If a file with the same name exists, REPLACE it.
    Returns destination Path.
    """
    downloads = _downloads_dir()
    target = _target_dir()

    src = _wait_for_download_complete(downloads, timeout=180, quiet_stabilize=1.0)

    sd = _fmt_date_for_name(start_date)
    ed = _fmt_date_for_name(end_date)
    base = f"Trapeze_{sd}_{ed}"
    dest = (target / base).with_suffix(src.suffix)

    try:
        # If a same-named file already exists, replace it
        if dest.exists():
            # On Windows, unlink will fail if the file is open. We surface a clear error in that case.
            dest.unlink()
        # Move into place
        shutil.move(str(src), str(dest))
    except PermissionError as e:
        raise PermissionError(
            f"Cannot replace '{dest}'. Is it open in Excel or another app?"
        ) from e

    print(f"Moved & replaced:\n  {src}\n→ {dest}")
    return dest

def previous_full_week_monday_sunday(today=None):
    """Return (last_monday, last_sunday) as MM/DD/YYYY strings."""
    if today is None:
        today = datetime.today()
    this_monday = today - timedelta(days=today.weekday())   # Monday of this week
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)
    return last_monday.strftime("%m/%d/%Y"), last_sunday.strftime("%m/%d/%Y")

# ---------- helpers ----------
def smart_click(driver, el):
    try:
        el.click(); return True
    except Exception:
        pass
    try:
        ActionChains(driver).move_to_element(el).pause(0.15).click(el).perform(); return True
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        driver.execute_script("arguments[0].click();", el); return True
    except Exception:
        return False

def switch_default(driver):
    try: driver.switch_to.default_content()
    except Exception: pass

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

def find_any_frame(driver, by, value, timeout=12):
    """Find ONE element across top + nested iframes (current window). Keeps focus where found."""
    end = time.time() + timeout
    while time.time() < end:
        for fr in all_iframes_recursive(driver):
            try:
                el = WebDriverWait(driver, 0.8).until(EC.presence_of_element_located((by, value)))
                return el, fr
            except Exception:
                continue
    raise TimeoutException(f"Element not found: {by}={value}")

def click_any_frame(driver, by, value, timeout=12):
    el, _ = find_any_frame(driver, by, value, timeout)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    if not smart_click(driver, el):
        raise TimeoutException(f"Could not click: {by}={value}")
    return True

def find_date_inputs(driver, timeout=15):
    """
    Return (start_input, end_input) for <input type='text' size='11'> fields.
    Picks the first two *visible* ones in whatever frame they live in.
    """
    end = time.time() + timeout
    xps = [
        "//input[@type='text' and @size='11' and @autocomplete='off']",
        "//input[@type='text' and @size='11']",
    ]
    while time.time() < end:
        for fr in all_iframes_recursive(driver):
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
            # de-dup
            uniq = []
            seen = set()
            for e in visibles:
                key = getattr(e, "id", None)
                if key and key in seen: continue
                seen.add(key); uniq.append(e)
            if len(uniq) >= 2:
                return uniq[0], uniq[1]
        time.sleep(0.2)
    raise TimeoutException("Could not find two visible date inputs (type=text, size=11).")

def set_input_value(driver, el, value):
    try:
        # Hard clear: Ctrl+A, Backspace
        el.click()
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.BACKSPACE)
        # Type value
        el.send_keys(value)
        # Fire common events
        driver.execute_script("arguments[0].dispatchEvent(new Event('input',{bubbles:true}))", el)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change',{bubbles:true}))", el)
        driver.execute_script("arguments[0].blur && arguments[0].blur();", el)
        return True
    except Exception:
        # JS fallback
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

def wait_new_window(driver, prev_handles, timeout=10):
    end = time.time() + timeout
    while time.time() < end:
        if len(driver.window_handles) > len(prev_handles):
            return True
        time.sleep(0.2)
    return False

def click_generated_report_link(driver, timeout=90):
    """
    After clicking Report, wait for the 'Please click here to obtain generated report.' link and click it.
    Searches across iframes; href is dynamic.
    """
    link_xpaths = [
        "//a[normalize-space(.)='Please click here to obtain generated report.']",
        "//a[contains(normalize-space(.), 'Please click here') and contains(normalize-space(.), 'obtain generated report')]",
    ]
    end = time.time() + timeout
    while time.time() < end:
        for fr in all_iframes_recursive(driver):
            for xp in link_xpaths:
                try:
                    el = WebDriverWait(driver, 0.8).until(EC.presence_of_element_located((By.XPATH, xp)))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    if smart_click(driver, el):
                        return True
                except Exception:
                    continue
        time.sleep(0.4)
    raise TimeoutException("Generated report link did not appear in time.")

# ----------------- core runner -----------------
def _run_once(url, username, password, start_date, end_date, download_dir=None, headless=False, keep_open=False):
    """
    Execute the browser flow for the given date range.
    Returns a summary string on success.
    Closes the browser at the end (even if an error occurs), unless keep_open=True.
    """
    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    # only keep browser open if you explicitly ask
    if keep_open:
        opts.add_experimental_option("detach", True)

    if download_dir:
        opts.add_experimental_option("prefs", {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    try:
        wait = WebDriverWait(driver, 30)
        driver.get(url)

        # --- Login ---
        wait.until(EC.element_to_be_clickable((By.ID, "loginName"))).send_keys(username)
        wait.until(EC.element_to_be_clickable((By.ID, "loginPassword"))).send_keys(password)
        smart_click(driver, wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.BtnLogin[name='login']"))))

        time.sleep(2)  # small settle

        # --- Recent → Trip Export Report for CIMS ---
        click_any_frame(driver, By.XPATH, "//div[contains(@class,'WorkspaceFootItem')][normalize-space()='Recent']", timeout=20)

        prev_handles = set(driver.window_handles)
        click_any_frame(driver, By.XPATH, "//td[normalize-space()='Trip Export Report for CIMS']", timeout=20)
        # If opening the recent item spawns a new window, switch to it
        if wait_new_window(driver, prev_handles, timeout=5):
            new_handle = list(set(driver.window_handles) - prev_handles)[0]
            driver.switch_to.window(new_handle)

        # --- Fill Start/End Date (robust against re-render) ---
        for attempt in range(3):
            try:
                start_input, end_input = find_date_inputs(driver, timeout=20)

                ok1 = set_input_value(driver, start_input, start_date)
                if not ok1:
                    start_input, _ = find_date_inputs(driver, timeout=10)
                    if not set_input_value(driver, start_input, start_date):
                        raise TimeoutException("Failed to set Start Date")

                time.sleep(0.3)  # many UIs validate/re-render after first date

                # Re-find end input fresh to avoid stale handles
                _, end_input = find_date_inputs(driver, timeout=10)
                ok2 = set_input_value(driver, end_input, end_date)
                if not ok2:
                    _, end_input = find_date_inputs(driver, timeout=10)
                    if not set_input_value(driver, end_input, end_date):
                        raise TimeoutException("Failed to set End Date")

                break  # success
            except StaleElementReferenceException:
                if attempt == 2:
                    raise
                time.sleep(0.2)

        # --- Click Report ---
        report_xpaths = [
            "//button[contains(@class,'BarItem') and normalize-space()='Report']",
            "//button[contains(@onclick,'3000018')]"
        ]
        clicked = False
        for xp in report_xpaths:
            try:
                click_any_frame(driver, By.XPATH, xp, timeout=15)
                clicked = True
                break
            except Exception:
                continue
        if not clicked:
            raise TimeoutException("Couldn't click Report button.")

        # --- Wait for and click the generated report link ---
        click_generated_report_link(driver, timeout=90)

        # --- Move & rename the downloaded file (Downloads → OneDrive/Trapeze) ---
        try:
            move_and_rename_report(start_date, end_date)
        except Exception as e:
            print(f"WARNING: Download rename/move failed: {e}")

        msg = f"Downloaded: clicked the generated report link for {start_date} → {end_date}."
        print(msg)
        return msg

    finally:
        # Close the browser unless explicitly told to keep it open
        if not keep_open:
            try:
                driver.quit()
            except Exception:
                pass


# ----------------- public API -----------------
def run(
    url: str = URL,
    username: str | None = None,
    password: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    download_dir: str | None = None,
    headless: bool = False,
):
    if username is None or password is None:
        username, password = _get_trapeze_creds()

    if not start_date or not end_date:
        start_date, end_date = previous_full_week_monday_sunday()
    return _run_once(url, username, password, start_date, end_date, download_dir, headless)


# ----------------- CLI entry -----------------
def main():
    # Default: previous full week (Mon→Sun)
    start_date, end_date = previous_full_week_monday_sunday()
    run(start_date=start_date, end_date=end_date)

if __name__ == "__main__":
    main()
