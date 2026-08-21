import getpass
import json
import os
import socket
import subprocess
import time
import requests
from seleniumwire import webdriver
from seleniumwire.utils import decode
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ================= Configuration =================
TARGET_URL = "https://fasih-sm.bps.go.id/app"

# Selectors (Full XPaths)
INIT_LOGIN_BTN_SELECTOR = (By.XPATH, "/html/body/div[1]/div/div/div/div[1]/div/div/div/div/div/div[2]/a[1]/span")
USERNAME_SELECTOR = (By.XPATH, "/html/body/div/div[2]/div/div/div/div/form/div[1]/input")
PASSWORD_SELECTOR = (By.XPATH, "/html/body/div/div[2]/div/div/div/div/form/div[2]/input")
LOGIN_BTN_SELECTOR = (By.XPATH, "/html/body/div/div[2]/div/div/div/div/form/div[4]/input[2]")

OTP_INPUT_SELECTOR = (By.XPATH, "/html/body/div/div[2]/div/div/form[1]/div[1]/div[2]/input")
OTP_SUBMIT_BTN_SELECTOR = (By.XPATH, "/html/body/div/div[2]/div/div/form[1]/div[2]/div[2]/input[2]")

# Request capture filter for surveys datatable
SURVEY_DATATABLE_KEYWORD = "survey/api/v1/surveys/datatable"

# Remote debugging configuration
DEBUGGER_PORT = 9222
DEBUGGER_HOST = "127.0.0.1"
DEBUGGER_ADDRESS = f"{DEBUGGER_HOST}:{DEBUGGER_PORT}"
CHROME_PROFILE_DIR = r"C:\chrome_debug_profile"
CHROME_EXECUTABLE_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]
# =================================================


def is_port_in_use(port, host=DEBUGGER_HOST):
    """Checks if the remote debugging port is already listening."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex((host, port)) == 0


def launch_chrome_with_debugging():
    """Finds Chrome and launches it with remote debugging enabled if not already running."""
    if is_port_in_use(DEBUGGER_PORT):
        print(f"[*] Chrome debugger already active on {DEBUGGER_ADDRESS}.")
        return

    chrome_exe = next((p for p in CHROME_EXECUTABLE_PATHS if os.path.exists(p)), None)
    if not chrome_exe:
        raise FileNotFoundError("Chrome executable not found in standard paths.")

    cmd = [
        chrome_exe,
        f"--remote-debugging-port={DEBUGGER_PORT}",
        f"--user-data-dir={CHROME_PROFILE_DIR}",
    ]
    print(f"[*] Launching Chrome: {' '.join(cmd)}")
    subprocess.Popen(cmd)

    # Wait until debugging port becomes available
    for _ in range(15):
        if is_port_in_use(DEBUGGER_PORT):
            print("[+] Chrome debugging port is ready.")
            return
        time.sleep(1)

    raise TimeoutError("Timed out waiting for Chrome to start with remote debugging.")


def get_driver():
    """Ensures Chrome is running with debugging, enables CDP Network interception, then connects to it."""
    launch_chrome_with_debugging()

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", DEBUGGER_ADDRESS)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )

    # Enable Chrome DevTools Protocol Network monitoring
    try:
        driver.execute_cdp_cmd("Network.enable", {})
    except Exception as e:
        print(f"[!] Warning: Could not enable CDP network tracking: {e}")

    return driver


def login_and_verify_otp(driver, url, username, password, otp=""):
    """Navigates to the website, clicks login, enters credentials, and handles OTP if provided."""
    print(f"[*] Navigating to: {url}")
    driver.get(url)

    wait = WebDriverWait(driver, 60)

    print("[*] Waiting for initial login button...")
    init_login_btn = wait.until(EC.element_to_be_clickable(INIT_LOGIN_BTN_SELECTOR))
    init_login_btn.click()

    print("[*] Waiting for username element...")
    user_input = wait.until(EC.visibility_of_element_located(USERNAME_SELECTOR))
    user_input.clear()
    user_input.send_keys(username)

    print("[*] Entering password...")
    pass_input = wait.until(EC.visibility_of_element_located(PASSWORD_SELECTOR))
    pass_input.clear()
    pass_input.send_keys(password)

    print("[*] Clicking login button...")
    login_btn = wait.until(EC.element_to_be_clickable(LOGIN_BTN_SELECTOR))
    login_btn.click()

    # If OTP was entered by the user in terminal
    if otp:
        print("[*] Waiting for OTP element...")
        otp_input = wait.until(EC.visibility_of_element_located(OTP_INPUT_SELECTOR))
        otp_input.clear()
        otp_input.send_keys(otp)

        print("[*] Submitting OTP...")
        otp_submit_btn = wait.until(EC.element_to_be_clickable(OTP_SUBMIT_BTN_SELECTOR))
        otp_submit_btn.click()
        print("[*] OTP submitted.")
    else:
        print("[*] No OTP entered. Proceeding to home page...")

    # Wait briefly for post-login page/requests to settle
    time.sleep(5)


def capture_survey_list(driver, survey_type="", timeout=30):
    """Captures cookies from the logged-in browser and uses requests to fetch all surveys."""
    print("\n[*] Extracting session cookies and requesting all surveys...")

    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # 1. Extract all cookies from the active browser session
            selenium_cookies = driver.get_cookies()
            cookies = {c["name"]: c["value"] for c in selenium_cookies}

            # 2. Extract XSRF-TOKEN
            xsrf_token = cookies.get("XSRF-TOKEN", "")

            # 3. Build headers exactly matching the browser request
            headers = {
                "accept": "*/*",
                "accept-language": "en-US,en;q=0.9,id;q=0.8",
                "content-type": "application/json",
                "origin": "https://fasih-sm.bps.go.id",
                "referer": "https://fasih-sm.bps.go.id/app/surveys?page=0&perPage=10&layout=list",
                "sec-ch-ua": '"Not=A?Brand";v="99", "Google Chrome";v="151", "Chromium";v="151"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
                "user-agent": driver.execute_script("return navigator.userAgent;"),
            }

            if xsrf_token:
                headers["x-xsrf-token"] = xsrf_token

            # 4. Check for Bearer token if stored in localStorage / sessionStorage
            try:
                token_val = driver.execute_script("""
                    for (let storage of [localStorage, sessionStorage]) {
                        for (let i = 0; i < storage.length; i++) {
                            let key = storage.key(i);
                            let val = storage.getItem(key);
                            if (!val) continue;
                            if (key.toLowerCase().includes('token') || key.toLowerCase().includes('auth')) {
                                try {
                                    let parsed = JSON.parse(val);
                                    if (typeof parsed === 'object' && parsed !== null) {
                                        val = parsed.token || parsed.accessToken || parsed.access_token || parsed.jwt || val;
                                    }
                                } catch (e) {}
                                if (typeof val === 'string' && val.length > 20) return val;
                            }
                        }
                    }
                    return null;
                """)
                if token_val:
                    headers["authorization"] = token_val if token_val.startswith("Bearer ") else f"Bearer {token_val}"
            except Exception:
                pass

            # 5. Send POST request via requests
            payload = {
                "pageNumber": 0,
                "pageSize": 100,
                "keywordSearch": "",
                "sortBy": "CREATED_AT",
                "sortDirection": "DESC",
            }

            resp = requests.post(
                "https://fasih-sm.bps.go.id/app/api/survey/api/v1/surveys/datatable",
                params={"surveyType": survey_type},
                cookies=cookies,
                headers=headers,
                json=payload,
                timeout=10,
            )

            if resp.status_code == 200:
                data = resp.json()
                surveys = data.get("data", {}).get("content", [])
                if surveys:
                    print(f"[+] Successfully fetched {len(surveys)} surveys.")
                    return surveys
            else:
                print(f"[!] Request returned status code: {resp.status_code}")

        except Exception as e:
            print(f"[!] Error: {e}")

        time.sleep(2)

    print("[!] Could not capture survey content.")
    return []


def prompt_user_select_survey(surveys):
    """Displays the list of surveys in terminal and prompts user to pick one."""
    if not surveys:
        print("[!] No surveys available to display.")
        return None

    print("\n" + "=" * 60)
    print("                    SURVEY LIST                    ")
    print("=" * 60)
    for idx, survey in enumerate(surveys, start=1):
        name = survey.get("name", "N/A")
        survey_type = survey.get("surveyType", "N/A")
        print(f"[{idx}] {name}")
        print(f"    Type: {survey_type}")

    print("=" * 60)
    while True:
        choice = input(f"Select a survey (1-{len(surveys)}): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(surveys):
            selected = surveys[int(choice) - 1]
            print(f"\n[+] Selected Survey: {selected.get('name')}")
            return selected
        print("[!] Invalid choice. Please enter a valid number.")


def fetch_survey_periods(driver, survey_id, timeout=30):
    """Fetches survey periods for the selected survey using the authenticated session."""
    print(f"\n[*] Fetching survey periods for survey ID: {survey_id}...")

    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            selenium_cookies = driver.get_cookies()
            cookies = {c["name"]: c["value"] for c in selenium_cookies}
            xsrf_token = cookies.get("XSRF-TOKEN", "")

            headers = {
                "accept": "*/*",
                "accept-language": "en-US,en;q=0.9",
                "referer": f"https://fasih-sm.bps.go.id/app/surveys/{survey_id}",
                "sec-ch-ua": '"Not=A?Brand";v="99", "Google Chrome";v="151", "Chromium";v="151"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
                "user-agent": driver.execute_script("return navigator.userAgent;"),
            }

            if xsrf_token:
                headers["x-xsrf-token"] = xsrf_token

            # Check Bearer token if stored in web storage
            try:
                token_val = driver.execute_script("""
                    for (let storage of [localStorage, sessionStorage]) {
                        for (let i = 0; i < storage.length; i++) {
                            let key = storage.key(i);
                            let val = storage.getItem(key);
                            if (!val) continue;
                            if (key.toLowerCase().includes('token') || key.toLowerCase().includes('auth')) {
                                try {
                                    let parsed = JSON.parse(val);
                                    if (typeof parsed === 'object' && parsed !== null) {
                                        val = parsed.token || parsed.accessToken || parsed.access_token || parsed.jwt || val;
                                    }
                                } catch (e) {}
                                if (typeof val === 'string' && val.length > 20) return val;
                            }
                        }
                    }
                    return null;
                """)
                if token_val:
                    headers["authorization"] = token_val if token_val.startswith("Bearer ") else f"Bearer {token_val}"
            except Exception:
                pass

            resp = requests.get(
                "https://fasih-sm.bps.go.id/app/api/survey/api/v1/survey-periods/my",
                params={"surveyId": survey_id},
                cookies=cookies,
                headers=headers,
                timeout=10,
            )

            if resp.status_code == 200:
                data = resp.json()
                periods = data.get("data", [])
                if periods:
                    print(f"[+] Successfully fetched {len(periods)} survey period(s).")
                    return periods
                else:
                    print("[!] No survey periods found in response data.")
                    return []
            else:
                print(f"[!] Request returned status code: {resp.status_code}")

        except Exception as e:
            print(f"[!] Error: {e}")

        time.sleep(2)

    print("[!] Could not retrieve survey periods.")
    return []


def prompt_user_select_periods(periods):
    """Displays period names and allows single or multi-selection (e.g. '1', '1,2', or 'all')."""
    if not periods:
        print("[!] No survey periods available to display.")
        return []

    print("\n" + "=" * 60)
    print("                 SURVEY PERIOD LIST                 ")
    print("=" * 60)
    for idx, period in enumerate(periods, start=1):
        name = period.get("name", "N/A")
        print(f"[{idx}] {name}")

    print("=" * 60)
    while True:
        choice = input(f"Select period(s) (e.g., 1 or 1,2 or all): ").strip()
        if not choice:
            print("[!] Please enter at least one choice.")
            continue

        if choice.lower() == "all":
            selected = periods
            print(f"\n[+] Selected all {len(selected)} period(s):")
            for p in selected:
                print(f"    - {p.get('name')}")
            return selected

        parts = [p.strip() for p in choice.split(",") if p.strip()]
        if parts and all(p.isdigit() and 1 <= int(p) <= len(periods) for p in parts):
            # deduplicate while maintaining order
            selected_indices = list(dict.fromkeys(int(p) - 1 for p in parts))
            selected = [periods[i] for i in selected_indices]
            print(f"\n[+] Selected {len(selected)} period(s):")
            for p in selected:
                print(f"    - {p.get('name')}")
            return selected

        print(f"[!] Invalid selection. Please enter numbers between 1 and {len(periods)} separated by commas, or 'all'.")


def main():
    print("=" * 50)
    print("      FASIH SCRAPER - LOGIN & CAPTURE       ")
    print("=" * 50)

    # Prompt user in terminal for credentials
    username = input("Enter Username / Email            : ").strip()
    password = getpass.getpass("Enter Password                    : ")
    otp = input("Enter OTP (leave blank if none)      : ").strip()

    driver = get_driver()
    try:
        # Step 1: Open website, perform login flow, submit OTP if provided
        login_and_verify_otp(driver, TARGET_URL, username, password, otp=otp)

        # Step 2: Capture all surveys (surveyType='')
        surveys = capture_survey_list(driver)

        # Step 3: Present surveys to user to select
        selected_survey = prompt_user_select_survey(surveys)

        # Step 4: Fetch and select survey period(s)
        if selected_survey and "id" in selected_survey:
            periods = fetch_survey_periods(driver, selected_survey["id"])
            selected_periods = prompt_user_select_periods(periods)

        # Keep browser open for inspection if needed
        input("\nPress Enter to close browser...")

    except Exception as e:
        print(f"[!] Error occurred: {e}")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
