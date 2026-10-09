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
from selenium.common.exceptions import (
    WebDriverException,
    TimeoutException,
    NoSuchElementException,
)

from webdriver_manager.chrome import ChromeDriverManager


# =========================================================
# CONFIG
# =========================================================
URL_LOGIN = "https://www.caredove.com/user/account"

URL_REPORTS_MAIN = (
    "https://www.caredove.com/organization/reports/orgid/3866"
)  # Reconnect Health Services

URL_REPORTS_WEST = (
    "https://www.caredove.com/organization/reports/orgid/3851"
)  # West Community Connect Hub


# Recommended:
# Store these as Windows environment variables
#
# PowerShell:
# setx CAREDOVE_EMAIL "your_email"
# setx CAREDOVE_PASSWORD "your_password"
#
EMAIL = os.getenv(
    "CAREDOVE_EMAIL",
    "yxu@reconnect.on.ca"
)

PASSWORD = os.getenv(
    "CAREDOVE_PASSWORD",
    "your_password"
)


DOWNLOAD_DIR = str(
    Path.home()
    / "Downloads"
    / "caredove"
)

DOWNLOAD_TIMEOUT = 120


# =========================================================
# DATE LOGIC
# Fiscal year starts April 1
# Pull current + previous fiscal year
# =========================================================
def last_two_fiscal_years_window(
    fy_start_month=4,
    fy_start_day=1
):
    """
    Returns:
        start_date, end_date

    Example:
    today = 2026-08-27

    current FY starts:
        2026-04-01

    report starts:
        2025-04-01

    report ends:
        today
    """

    today = date.today()

    current_fy_start_year = today.year

    if (
        today.month,
        today.day
    ) < (
        fy_start_month,
        fy_start_day
    ):
        current_fy_start_year -= 1

    start = date(
        current_fy_start_year - 1,
        fy_start_month,
        fy_start_day
    )

    end = today

    return (
        start.strftime("%Y-%m-%d"),
        end.strftime("%Y-%m-%d")
    )


# =========================================================
# HELPERS
# =========================================================
def wait_ready(
    driver,
    timeout=25
):
    WebDriverWait(
        driver,
        timeout
    ).until(
        lambda d:
        d.execute_script(
            "return document.readyState"
        )
        == "complete"
    )


def safe_click(
    driver,
    elem
):
    # Normal click
    try:
        elem.click()
        return True

    except WebDriverException:
        pass

    # Scroll then click
    try:
        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block:'center'
            });
            """,
            elem
        )

        time.sleep(0.2)

        elem.click()

        return True

    except WebDriverException:
        pass

    # ActionChains
    try:
        ActionChains(
            driver
        ).move_to_element(
            elem
        ).pause(
            0.1
        ).click().perform()

        return True

    except WebDriverException:
        pass

    # JavaScript click
    try:
        driver.execute_script(
            "arguments[0].click();",
            elem
        )

        return True

    except WebDriverException:
        return False


def snapshot_dir(
    download_dir: str
):
    return set(
        os.listdir(
            download_dir
        )
    )


def _size_stable(
    path: Path,
    checks=3,
    interval=0.5
):
    try:
        previous_size = (
            path.stat().st_size
        )

        for _ in range(checks):

            time.sleep(
                interval
            )

            current_size = (
                path.stat().st_size
            )

            if (
                current_size
                != previous_size
            ):
                previous_size = (
                    current_size
                )

                return False

        return True

    except FileNotFoundError:
        return False


def wait_for_new_file(
    download_dir: str,
    before_snapshot: set,
    timeout: int
) -> Path:

    start = time.time()

    download_path = Path(
        download_dir
    )

    while (
        time.time() - start
        < timeout
    ):

        after_snapshot = set(
            os.listdir(
                download_dir
            )
        )

        new_names = (
            after_snapshot
            - before_snapshot
        )

        if new_names:

            candidates = [
                download_path / name
                for name in new_names
                if not str(
                    name
                ).endswith(
                    ".crdownload"
                )
            ]

            if candidates:

                newest = max(
                    candidates,
                    key=lambda p:
                    p.stat().st_mtime
                )

                if _size_stable(
                    newest
                ):
                    return newest

        time.sleep(
            0.5
        )

    raise FileNotFoundError(
        "No completed download "
        "detected within timeout"
    )


def download_and_rename(
    driver,
    wait,
    download_btn_xpath: str,
    download_dir: str,
    new_filename: str
):
    before_snapshot = (
        snapshot_dir(
            download_dir
        )
    )

    download_button = (
        wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    download_btn_xpath
                )
            )
        )
    )

    if not safe_click(
        driver,
        download_button
    ):
        raise RuntimeError(
            "Could not click "
            "download button"
        )

    downloaded_file = (
        wait_for_new_file(
            download_dir,
            before_snapshot,
            timeout=DOWNLOAD_TIMEOUT
        )
    )

    target = (
        Path(download_dir)
        / f"{new_filename}.xlsx"
    )

    if target.exists():
        target.unlink()

    downloaded_file.rename(
        target
    )

    print(
        f"Saved → {target}"
    )


def open_report_dropdown(
    wait
):
    return wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//button["
                "@data-id='report_select' "
                "and contains("
                "@class,"
                "'dropdown-toggle'"
                ")]"
            )
        )
    )


def choose_dropdown_option(
    wait,
    driver,
    option_text: str
):
    button = (
        open_report_dropdown(
            wait
        )
    )

    safe_click(
        driver,
        button
    )

    option = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                (
                    "//span["
                    "@class='text' "
                    "and normalize-space()="
                    f"'{option_text}'"
                    "]/.."
                )
            )
        )
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block:'center'
        });
        """,
        option
    )

    if not safe_click(
        driver,
        option
    ):
        raise RuntimeError(
            f"Could not click option: "
            f"{option_text}"
        )


def set_custom_date_range(
    wait,
    driver
):
    """
    Select Custom date range and
    set start/end dates using
    the fiscal-year logic.
    """

    start_str, end_str = (
        last_two_fiscal_years_window()
    )

    # -----------------------------------------------------
    # Select custom date range
    # -----------------------------------------------------
    try:
        label = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "label[for='op6']"
                )
            )
        )

        safe_click(
            driver,
            label
        )

    except TimeoutException:
        pass

    # -----------------------------------------------------
    # Ensure radio is checked
    # -----------------------------------------------------
    try:
        custom_radio = (
            driver.find_element(
                By.ID,
                "op6"
            )
        )

        driver.execute_script(
            """
            if(arguments[0]){
                arguments[0].checked=true;

                arguments[0].dispatchEvent(
                    new Event(
                        'change',
                        {bubbles:true}
                    )
                );

                arguments[0].dispatchEvent(
                    new Event(
                        'click',
                        {bubbles:true}
                    )
                );
            }
            """,
            custom_radio
        )

    except NoSuchElementException:
        raise RuntimeError(
            "CUSTOM date range "
            "radio (#op6) not found"
        )

    # -----------------------------------------------------
    # Locate date inputs
    # -----------------------------------------------------
    start_input = None
    end_input = None

    try:
        start_input = wait.until(
            EC.visibility_of_element_located(
                (
                    By.ID,
                    "startDate"
                )
            )
        )

    except TimeoutException:
        pass

    try:
        end_input = wait.until(
            EC.visibility_of_element_located(
                (
                    By.ID,
                    "endDate"
                )
            )
        )

    except TimeoutException:
        pass

    # Fallback
    if (
        start_input is None
        or end_input is None
    ):

        inputs = wait.until(
            EC.presence_of_all_elements_located(
                (
                    By.CSS_SELECTOR,
                    "input.rangeDates"
                )
            )
        )

        if inputs:

            if start_input is None:
                start_input = (
                    inputs[0]
                )

            if end_input is None:
                end_input = (
                    inputs[-1]
                    if len(inputs) > 1
                    else inputs[0]
                )

    if (
        start_input is None
        or end_input is None
    ):
        raise RuntimeError(
            "Could not locate "
            "start/end date inputs"
        )

    # -----------------------------------------------------
    # Set values
    # -----------------------------------------------------
    for input_element, value in (
        (
            start_input,
            start_str
        ),
        (
            end_input,
            end_str
        )
    ):

        driver.execute_script(
            """
            arguments[0]
                .removeAttribute(
                    'readonly'
                );

            arguments[0]
                .removeAttribute(
                    'disabled'
                );
            """,
            input_element
        )

        driver.execute_script(
            """
            arguments[0].value =
                arguments[1];

            arguments[0].dispatchEvent(
                new Event(
                    'input',
                    {bubbles:true}
                )
            );

            arguments[0].dispatchEvent(
                new Event(
                    'change',
                    {bubbles:true}
                )
            );
            """,
            input_element,
            value
        )

    print(
        f"Set date range: "
        f"{start_str} → {end_str}"
    )


# =========================================================
# LOGIN
# =========================================================
def login(
    driver,
    wait
):
    """
    New Caredove login flow:

    1. Enter email
    2. Click Continue
    3. Enter password
    4. Click Sign In
    """

    print(
        "\nLogging into Caredove..."
    )

    driver.get(
        URL_LOGIN
    )

    wait_ready(
        driver
    )

    # -----------------------------------------------------
    # STEP 1:
    # Enter email
    #
    # <input
    #   id="EmailFirst_email"
    #   name="EmailFirst[email]"
    # >
    # -----------------------------------------------------
    email_input = wait.until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "EmailFirst_email"
            )
        )
    )

    email_input.clear()

    email_input.send_keys(
        EMAIL
    )

    print(
        "Email entered."
    )

    # -----------------------------------------------------
    # STEP 2:
    # Click Continue
    #
    # <button
    #   id="emailFirstSubmit"
    # >
    # -----------------------------------------------------
    continue_button = wait.until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "emailFirstSubmit"
            )
        )
    )

    if not safe_click(
        driver,
        continue_button
    ):
        raise RuntimeError(
            "Could not click "
            "Continue button"
        )

    print(
        "Continue clicked."
    )

    # -----------------------------------------------------
    # STEP 3:
    # Wait for password page
    #
    # <input
    #   id="LoginForm_u_password"
    # >
    # -----------------------------------------------------
    password_input = wait.until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "LoginForm_u_password"
            )
        )
    )

    password_input.clear()

    password_input.send_keys(
        PASSWORD
    )

    print(
        "Password entered."
    )

    # -----------------------------------------------------
    # STEP 4:
    # Click Sign In
    #
    # <button
    #   id="formSubmit"
    # >
    # -----------------------------------------------------
    sign_in_button = wait.until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "formSubmit"
            )
        )
    )

    if not safe_click(
        driver,
        sign_in_button
    ):
        raise RuntimeError(
            "Could not click "
            "Sign In button"
        )

    print(
        "Sign In clicked."
    )

    # Wait for login / redirect
    wait_ready(
        driver
    )

    # Give dynamic page a moment
    time.sleep(
        1
    )

    print(
        "Login completed."
    )


# =========================================================
# MAIN DOWNLOAD FLOW
# =========================================================
def run_and_download():

    os.makedirs(
        DOWNLOAD_DIR,
        exist_ok=True
    )

    # =====================================================
    # Chrome configuration
    # =====================================================
    options = Options()

    options.add_argument(
        "--start-maximized"
    )

    preferences = {
        "download.default_directory":
            DOWNLOAD_DIR,

        "download.prompt_for_download":
            False,

        "download.directory_upgrade":
            True,

        "safebrowsing.enabled":
            True,
    }

    options.add_experimental_option(
        "prefs",
        preferences
    )

    # =====================================================
    # Start Chrome
    # =====================================================
    driver = webdriver.Chrome(
        service=Service(
            ChromeDriverManager()
            .install()
        ),
        options=options
    )

    wait = WebDriverWait(
        driver,
        30
    )

    try:

        # =================================================
        # LOGIN
        # =================================================
        login(
            driver,
            wait
        )

        # =================================================
        # Go to Reconnect Reports
        # orgid = 3866
        # =================================================
        try:
            reports_link = wait.until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        (
                            "//a["
                            "@href="
                            "'/organization/reports/"
                            "orgid/3866'"
                            "]"
                        )
                    )
                )
            )

            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        (
                            "//a["
                            "@href="
                            "'/organization/reports/"
                            "orgid/3866'"
                            "]"
                        )
                    )
                )
            )

            if not safe_click(
                driver,
                reports_link
            ):
                raise WebDriverException(
                    "Couldn't click "
                    "reports link"
                )

        except (
            TimeoutException,
            WebDriverException
        ):

            driver.get(
                URL_REPORTS_MAIN
            )

        wait_ready(
            driver
        )

        # =================================================
        # Download button
        # =================================================
        download_btn_xpath = (
            "//button["
            "contains(@class,'download_report') "
            "and contains("
            ".,"
            "'Download Data (.xlsx)'"
            ")"
            "]"
        )

        # =================================================
        # RECONNECT - INBOUND
        # =================================================
        print(
            "\nDownloading "
            "Reconnect Inbound..."
        )

        choose_dropdown_option(
            wait,
            driver,
            (
                "Spreadsheet: "
                "Inbound Referral Raw Data"
            )
        )

        set_custom_date_range(
            wait,
            driver
        )

        download_and_rename(
            driver,
            wait,
            download_btn_xpath,
            DOWNLOAD_DIR,
            "CaredoveBookingsReport_Reconnect"
        )

        # =================================================
        # RECONNECT - OUTBOUND
        # =================================================
        print(
            "\nDownloading "
            "Reconnect Outbound..."
        )

        choose_dropdown_option(
            wait,
            driver,
            (
                "Spreadsheet: "
                "Outbound Referral Raw Data"
            )
        )

        set_custom_date_range(
            wait,
            driver
        )

        download_and_rename(
            driver,
            wait,
            download_btn_xpath,
            DOWNLOAD_DIR,
            (
                "CaredoveBookingsReport_"
                "Outbound_Reconnect"
            )
        )

        # =================================================
        # Navigate to West Hub
        # orgid = 3851
        # =================================================
        driver.get(
            URL_REPORTS_WEST
        )

        wait_ready(
            driver
        )

        # =================================================
        # WEST - INBOUND
        # =================================================
        print(
            "\nDownloading "
            "West Inbound..."
        )

        choose_dropdown_option(
            wait,
            driver,
            (
                "Spreadsheet: "
                "Inbound Referral Raw Data"
            )
        )

        set_custom_date_range(
            wait,
            driver
        )

        download_and_rename(
            driver,
            wait,
            download_btn_xpath,
            DOWNLOAD_DIR,
            "CaredoveBookingsReport_West"
        )

        # =================================================
        # WEST - OUTBOUND
        # =================================================
        print(
            "\nDownloading "
            "West Outbound..."
        )

        choose_dropdown_option(
            wait,
            driver,
            (
                "Spreadsheet: "
                "Outbound Referral Raw Data"
            )
        )

        set_custom_date_range(
            wait,
            driver
        )

        download_and_rename(
            driver,
            wait,
            download_btn_xpath,
            DOWNLOAD_DIR,
            (
                "CaredoveBookingsReport_"
                "Outbound_West"
            )
        )

        print(
            "\n========================================="
        )

        print(
            "All four reports downloaded "
            "and renamed ✅"
        )

        print(
            "========================================="
        )

    finally:

        driver.quit()


# =========================================================
# RUN FUNCTION
# =========================================================
def run():
    run_and_download()


# =========================================================
# RUN DIRECTLY
# =========================================================
if __name__ == "__main__":
    run()