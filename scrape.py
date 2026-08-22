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
from extract_excel import extract_json_to_excel

# ================= Configuration =================
TARGET_URL = "https://fasih-sm.bps.go.id/app"

# Global delays & rate limit configuration (in seconds)
DEFAULT_DELAY = 5  # Delay between iterative requests
DETAIL_DELAY = 5   # Delay between fetching respondent detail data

# Region configuration
ENABLE_USER_PROMPT_REGION_LEVEL = True  # Set True to prompt user for target level, False to use maximum region level from metadata
ROOT_REGION_NAME = "JAWA TIMUR"         # Root region name to start building hierarchy
ROOT_REGION_CODE = "35"                 # Root region code

# Output configuration
RESULT_DIR = "result"                   # Folder to save output Excel files
OUTPUT_FILENAME = "data.xlsx"          # Output file name

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


def get_session_auth(driver, referer="https://fasih-sm.bps.go.id"):
    """Extracts session cookies and headers (including XSRF & Bearer token) from the active browser."""
    selenium_cookies = driver.get_cookies()
    cookies = {c["name"]: c["value"] for c in selenium_cookies}
    xsrf_token = cookies.get("XSRF-TOKEN", "")

    headers = {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9,id;q=0.8",
        "origin": "https://fasih-sm.bps.go.id",
        "referer": referer,
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

    return cookies, headers


def capture_survey_list(driver, survey_type="", timeout=30):
    """Captures cookies from the logged-in browser and uses requests to fetch all surveys."""
    print("\n[*] Extracting session cookies and requesting all surveys...")

    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            cookies, headers = get_session_auth(driver, referer="https://fasih-sm.bps.go.id/app/surveys?page=0&perPage=10&layout=list")
            headers["content-type"] = "application/json"

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
            cookies, headers = get_session_auth(driver, referer=f"https://fasih-sm.bps.go.id/app/surveys/{survey_id}")

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


def prompt_user_select_mode():
    """Displays action mode prompt immediately after credentials."""
    print("\n" + "=" * 60)
    print("                    ACTION MODE                    ")
    print("=" * 60)
    print("[1] Start a new download (delete existing data.xlsx and json results)")
    print("[2] Resume download (skip completed regions & existing JSON details)")
    print("[3] Extract only (generate Excel from existing JSON files without downloading)")
    print("=" * 60)
    while True:
        choice = input("Select action mode (1, 2, or 3): ").strip()
        if choice in ("1", "2", "3"):
            selected_mode = int(choice)
            mode_names = {
                1: "Start New Download",
                2: "Resume Download",
                3: "Extract Only"
            }
            print(f"\n[+] Selected Mode: [{selected_mode}] {mode_names[selected_mode]}")
            return selected_mode
        print("[!] Invalid choice. Please enter 1, 2, or 3.")


def get_progress_filepath(survey_id, period_id):
    """Returns path to the region progress JSON file for a period."""
    return os.path.join(RESULT_DIR, str(survey_id), str(period_id), "region_progress.json")


def load_region_progress(survey_id, period_id):
    """Loads region completion tracking for a period."""
    progress_file = get_progress_filepath(survey_id, period_id)
    if os.path.exists(progress_file):
        try:
            with open(progress_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"completed_regions": []}


def save_region_progress(survey_id, period_id, progress):
    """Saves region completion tracking for a period."""
    progress_file = get_progress_filepath(survey_id, period_id)
    os.makedirs(os.path.dirname(progress_file), exist_ok=True)
    try:
        with open(progress_file, "w", encoding="utf-8") as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[!] Error saving region progress: {e}")


def clear_period_data(survey_id, period_id):
    """Deletes data.xlsx and json directory for a specific survey period."""
    import shutil
    period_dir = os.path.join(RESULT_DIR, str(survey_id), str(period_id))
    json_dir = os.path.join(period_dir, "json")
    excel_file = os.path.join(period_dir, OUTPUT_FILENAME)
    progress_file = get_progress_filepath(survey_id, period_id)

    if os.path.exists(excel_file):
        try:
            os.remove(excel_file)
            print(f"[+] Deleted existing Excel file: {excel_file}")
        except Exception as e:
            print(f"[!] Warning: Could not delete {excel_file}: {e}")

    if os.path.exists(json_dir):
        try:
            shutil.rmtree(json_dir)
            print(f"[+] Deleted existing JSON directory: {json_dir}")
        except Exception as e:
            print(f"[!] Warning: Could not delete {json_dir}: {e}")

    if os.path.exists(progress_file):
        try:
            os.remove(progress_file)
            print(f"[+] Reset region progress tracking: {progress_file}")
        except Exception as e:
            print(f"[!] Warning: Could not delete {progress_file}: {e}")



def fetch_survey_metadata(driver, survey_id, timeout=30):
    """Fetches survey details to obtain regionGroupId."""
    print(f"\n[*] Fetching metadata for survey ID: {survey_id}...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            cookies, headers = get_session_auth(driver, referer=f"https://fasih-sm.bps.go.id/app/surveys/{survey_id}")
            resp = requests.get(
                f"https://fasih-sm.bps.go.id/app/api/survey/api/v1/surveys/{survey_id}",
                cookies=cookies,
                headers=headers,
                timeout=10,
            )
            if resp.status_code == 200:
                data = resp.json().get("data", {})
                region_group_id = (
                    data.get("regionGroupId")
                    or data.get("regionGroup", {}).get("id")
                    or data.get("regionMetadataId")
                    or data.get("groupId")
                )
                if region_group_id:
                    print(f"[+] Region Group ID: {region_group_id}")
                    return region_group_id, data
                else:
                    print(f"[!] Warning: regionGroupId not found in survey response: {list(data.keys())}")
                    return None, data
            else:
                print(f"[!] Survey metadata request returned status code: {resp.status_code}")
        except Exception as e:
            print(f"[!] Error fetching survey metadata: {e}")
        time.sleep(2)
    return None, {}


def fetch_region_metadata(driver, region_group_id, timeout=30):
    """Fetches region metadata levels for the given region group ID."""
    print(f"\n[*] Fetching region metadata for Group ID: {region_group_id}...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            cookies, headers = get_session_auth(driver)
            resp = requests.get(
                "https://fasih-sm.bps.go.id/app/api/region/api/v1/region-metadata",
                params={"id": region_group_id},
                cookies=cookies,
                headers=headers,
                timeout=10,
            )
            if resp.status_code == 200:
                metadata = resp.json().get("data", {})
                levels = metadata.get("level", [])
                print(f"[+] Region metadata loaded. Total levels: {len(levels)}")
                return metadata
            else:
                print(f"[!] Region metadata returned status code: {resp.status_code}")
        except Exception as e:
            print(f"[!] Error fetching region metadata: {e}")
        time.sleep(2)
    return {}


def prompt_user_select_region_level(region_metadata):
    """Displays available region levels and prompts user to pick a target level."""
    levels = region_metadata.get("level", [])
    if not levels:
        print("[!] No region levels defined. Defaulting to Level 1.")
        return 1

    if not ENABLE_USER_PROMPT_REGION_LEVEL:
        target_level = region_metadata.get("levelCount") or len(levels)
        selected_level_name = levels[target_level - 1].get("name", f"Level {target_level}")
        print(f"[*] Region level prompt disabled. Using maximum region level: [{target_level}] {selected_level_name}")
        return target_level

    print("\n" + "=" * 60)
    print("                    REGION LEVELS                   ")
    print("=" * 60)
    for lvl in levels:
        lvl_id = lvl.get("id")
        lvl_name = lvl.get("name", "N/A")
        print(f"[{lvl_id}] {lvl_name}")

    print("=" * 60)
    while True:
        choice = input(f"Select target region level (1-{len(levels)}): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(levels):
            target_level = int(choice)
            selected_level_name = levels[target_level - 1].get("name", f"Level {target_level}")
            print(f"\n[+] Selected Target Level: [{target_level}] {selected_level_name}")
            return target_level
        print(f"[!] Invalid choice. Please enter a number between 1 and {len(levels)}.")


def fetch_regions_for_level(driver, group_id, level_num, parent_full_code=None):
    """Fetches regions at a specific level."""
    cookies, headers = get_session_auth(driver)
    url = f"https://fasih-sm.bps.go.id/app/api/region/api/v1/region/level{level_num}"
    params = {"groupId": group_id}
    if level_num > 1 and parent_full_code:
        params[f"level{level_num - 1}FullCode"] = parent_full_code

    try:
        resp = requests.get(url, params=params, cookies=cookies, headers=headers, timeout=10)
        if resp.status_code == 200:
            return resp.json().get("data", [])
        else:
            print(f"[!] Level {level_num} request failed with status: {resp.status_code}")
    except Exception as e:
        print(f"[!] Error fetching regions at level {level_num}: {e}")
    return []


def build_region_hierarchy(driver, group_id, target_level):
    """Traverses regions starting from root region down to target_level and returns list of paths."""
    print(f"\n[*] Fetching Level 1 regions...")
    level1_regions = fetch_regions_for_level(driver, group_id, 1)

    # Locate root region
    jt_region = next(
        (r for r in level1_regions if r.get("name", "").strip().upper() == ROOT_REGION_NAME.upper() or r.get("fullCode") == ROOT_REGION_CODE),
        None,
    )
    if not jt_region:
        print(f"[!] Warning: '{ROOT_REGION_NAME}' not found in Level 1 regions. Available provinces:")
        for r in level1_regions:
            print(f"    - {r.get('name')} (code: {r.get('fullCode')})")
        if level1_regions:
            jt_region = level1_regions[0]
            print(f"[*] Falling back to: {jt_region.get('name')}")
        else:
            return []

    print(f"[+] Root region selected: {jt_region.get('name')} (Code: {jt_region.get('fullCode')})")

    # List of lists representing region branch paths from Level 1 down to current level
    current_paths = [[jt_region]]

    for lvl in range(2, target_level + 1):
        print(f"[*] Fetching Level {lvl} sub-regions for {len(current_paths)} parent region(s)...")
        next_paths = []
        for path in current_paths:
            parent = path[-1]
            children = fetch_regions_for_level(driver, group_id, lvl, parent_full_code=parent.get("fullCode"))
            for child in children:
                next_paths.append(path + [child])
        current_paths = next_paths
        print(f"[+] Total region paths at Level {lvl}: {len(current_paths)}")

    return current_paths


def fetch_survey_respondents(driver, survey_period_id, survey_id, region_path=None, delay=DEFAULT_DELAY):
    """Fetches all respondents for a survey period and region path by paginating through results."""
    region_desc = " > ".join([r.get("name", "") for r in region_path]) if region_path else f"Period {survey_period_id}"
    print(f"\n[*] Fetching respondents for: {region_desc}...")

    cookies, headers = get_session_auth(
        driver,
        referer=f"https://fasih-sm.bps.go.id/app/surveys/{survey_id}/{survey_period_id}/data?page=1&perPage=100",
    )
    headers["content-type"] = "application/json"

    endpoint = "https://fasih-sm.bps.go.id/app/api/analytic/api/v2/assignment/datatable-all-user-survey-periode"
    all_respondents = []
    start = 0
    page_size = 100
    total_hit = None

    assignment_extra_param = {
        "surveyPeriodId": survey_period_id,
        "assignmentErrorStatusType": -1,
        "filterTargetType": "TARGET_ONLY",
    }
    if region_path:
        for idx, reg in enumerate(region_path, start=1):
            assignment_extra_param[f"region{idx}Id"] = reg["id"]

    while True:
        payload = {
            "start": start,
            "length": page_size,
            "columns": [
                {"data": "id", "orderable": True},
                {"data": "codeIdentity", "orderable": True},
                {"data": "data1", "orderable": True},
                {"data": "data2", "orderable": True},
                {"data": "data3", "orderable": True},
                {"data": "data4", "orderable": True},
                {"data": "data5", "orderable": True},
                {"data": "data6", "orderable": True},
                {"data": "data7", "orderable": True},
                {"data": "data8", "orderable": True},
                {"data": "data9", "orderable": True},
                {"data": "data10", "orderable": True},
            ],
            "order": [],
            "search": {"value": "", "regex": False},
            "assignmentExtraParam": assignment_extra_param,
        }

        try:
            resp = requests.post(
                endpoint,
                cookies=cookies,
                headers=headers,
                json=payload,
                timeout=15,
            )

            if resp.status_code != 200:
                print(f"[!] Error at start={start}: Status {resp.status_code}")
                break

            data = resp.json()
            if total_hit is None:
                total_hit = data.get("totalHit", 0)
                print(f"[+] Total respondents: {total_hit}")

            if total_hit == 0:
                break

            batch = data.get("searchData", [])
            if not batch:
                break

            all_respondents.extend(batch)
            print(f"[*] Fetched {len(all_respondents)}/{total_hit} respondents")

            if len(all_respondents) >= total_hit or len(batch) < page_size:
                break

            start += page_size
            time.sleep(delay)

        except Exception as e:
            print(f"[!] Exception at start={start}: {e}")
            break

    print(f"[+] Retrieved {len(all_respondents)} respondents for [{region_desc}].")
    return all_respondents


def fetch_assignment_detail(driver, assignment_id, period_id, timeout=15):
    """Fetches full assignment details for a given assignment ID."""
    cookies, headers = get_session_auth(
        driver,
        referer=f"https://fasih-sm.bps.go.id/app/assignment/{period_id}/{assignment_id}",
    )
    url = "https://fasih-sm.bps.go.id/app/api/assignment-general/api/assignment/get-by-assignment-id"
    params = {"assignmentId": assignment_id}

    try:
        resp = requests.get(url, params=params, cookies=cookies, headers=headers, timeout=timeout)
        if resp.status_code == 200:
            return resp.json().get("data", {})
        elif resp.status_code == 429:
            print(f"[!] Rate limited (429). Pausing for 10 seconds before retrying assignment {assignment_id}...")
            time.sleep(10)
            resp = requests.get(url, params=params, cookies=cookies, headers=headers, timeout=timeout)
            if resp.status_code == 200:
                return resp.json().get("data", {})
            else:
                print(f"[!] Retry for assignment {assignment_id} failed with status: {resp.status_code}")
        else:
            print(f"[!] Assignment detail request failed with status: {resp.status_code}")
    except Exception as e:
        print(f"[!] Exception fetching detail for assignment {assignment_id}: {e}")
    return {}


def enrich_respondents_with_details(driver, respondents, survey_id, period_id, delay=DETAIL_DELAY):
    """Fetches details for each respondent, caches to result/{survey_id}/{period_id}/json/{id}.json, and loads existing JSON if present."""
    if not respondents:
        return respondents

    json_dir = os.path.join(RESULT_DIR, str(survey_id), str(period_id), "json")
    os.makedirs(json_dir, exist_ok=True)

    print(f"\n[*] Processing details for {len(respondents)} respondent(s)...")
    for idx, item in enumerate(respondents, start=1):
        assignment_id = item.get("id")
        if not assignment_id:
            continue

        json_path = os.path.join(json_dir, f"{assignment_id}.json")

        if os.path.exists(json_path):
            print(f"[*] [{idx}/{len(respondents)}] Loading cached detail from JSON: {json_path}")
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    detail = json.load(f)
                item["assignment_detail"] = detail
                item["detail_data_raw"] = detail.get("data")
                item["detail_predefined_raw"] = detail.get("pre_defined_data")
                continue
            except Exception as e:
                print(f"[!] Error reading cached JSON for {assignment_id}: {e}. Re-fetching...")

        print(f"[*] [{idx}/{len(respondents)}] Fetching detail for assignment ID: {assignment_id}...")
        detail = fetch_assignment_detail(driver, assignment_id, period_id)
        if detail:
            item["assignment_detail"] = detail
            item["detail_data_raw"] = detail.get("data")
            item["detail_predefined_raw"] = detail.get("pre_defined_data")

            try:
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(detail, f, ensure_ascii=False, indent=2)
                print(f"[+] Saved JSON detail: {json_path}")
            except Exception as e:
                print(f"[!] Error saving JSON for {assignment_id}: {e}")

        time.sleep(delay)

    return respondents


def save_respondents_to_excel(respondents, survey_id, period_id):
    """Saves collected respondents to {RESULT_DIR}/{survey_id}/{period_id}/{OUTPUT_FILENAME}."""
    if not respondents:
        print(f"[!] No respondents to save for period {period_id}.")
        return

    output_dir = os.path.join(RESULT_DIR, str(survey_id), str(period_id))
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, OUTPUT_FILENAME)

    flattened_data = []
    for item in respondents:
        row = {}
        for k, v in item.items():
            if isinstance(v, (dict, list)):
                row[k] = json.dumps(v, ensure_ascii=False)
            else:
                row[k] = v
        flattened_data.append(row)

    try:
        import pandas as pd
        df = pd.DataFrame(flattened_data)
        df.to_excel(filepath, index=False)
        print(f"[+] Saved {len(respondents)} respondents to: {filepath}")
    except ImportError:
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Data"
            if flattened_data:
                headers = list(flattened_data[0].keys())
                ws.append(headers)
                for item in flattened_data:
                    ws.append([item.get(h, "") for h in headers])
            wb.save(filepath)
            print(f"[+] Saved {len(respondents)} respondents to: {filepath}")
        except Exception as e:
            print(f"[!] Error saving Excel file with openpyxl: {e}")
    except Exception as e:
        print(f"[!] Error saving Excel file: {e}")


def main():
    print("=" * 50)
    print("      FASIH SCRAPER - LOGIN & CAPTURE       ")
    print("=" * 50)

    # Prompt user in terminal for credentials
    username = input("Enter Username / Email            : ").strip()
    password = getpass.getpass("Enter Password                    : ")
    otp = input("Enter OTP (leave blank if none)      : ").strip()

    # Mode option asked right after credential prompt
    action_mode = prompt_user_select_mode()

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
            survey_id = selected_survey["id"]
            periods = fetch_survey_periods(driver, survey_id)
            selected_periods = prompt_user_select_periods(periods)

            if action_mode == 3:
                # Mode 3: Extract only from existing downloaded JSON files
                for period in selected_periods:
                    period_id = period.get("id")
                    period_name = period.get("name", "N/A")
                    print("\n" + "=" * 60)
                    print(f"[*] Extracting Excel for Period: {period_name}")
                    print("=" * 60)
                    extract_json_to_excel(survey_id, period_id)
            else:
                # Modes 1 & 2: Downloading data
                # Step 5: Fetch region metadata and prompt target level
                region_group_id, _ = fetch_survey_metadata(driver, survey_id)
                if region_group_id:
                    region_metadata = fetch_region_metadata(driver, region_group_id)
                    target_level = prompt_user_select_region_level(region_metadata)

                    # Step 6: Build region hierarchy starting from JAWA TIMUR down to target_level
                    region_paths = build_region_hierarchy(driver, region_group_id, target_level)
                    print(f"\n[+] Total target regions to scrape: {len(region_paths)}")

                    if selected_periods and region_paths:
                        for period in selected_periods:
                            period_id = period.get("id")
                            period_name = period.get("name", "N/A")
                            print("\n" + "=" * 60)
                            print(f"[*] Processing Period: {period_name}")
                            print("=" * 60)

                            if action_mode == 1:
                                print(f"[*] Mode 1 (New Download): Clearing existing data for Period '{period_name}'...")
                                clear_period_data(survey_id, period_id)

                            progress = load_region_progress(survey_id, period_id)
                            completed_region_keys = set(progress.get("completed_regions", []))

                            # Dynamically evaluate if all target regions are completed
                            all_target_keys = {" > ".join([r.get("fullCode", r.get("id", "")) for r in p]) for p in region_paths} if region_paths else set()
                            all_regions_completed = all_target_keys.issubset(completed_region_keys) if all_target_keys else False

                            # Load existing respondents from data.xlsx or cached JSONs if resuming
                            existing_respondents_by_id = {}
                            excel_path = os.path.join(RESULT_DIR, str(survey_id), str(period_id), OUTPUT_FILENAME)
                            if os.path.exists(excel_path):
                                try:
                                    import pandas as pd
                                    df_existing = pd.read_excel(excel_path)
                                    id_col = "assignment_id" if "assignment_id" in df_existing.columns else ("id" if "id" in df_existing.columns else None)
                                    if id_col:
                                        records = df_existing.to_dict("records")
                                        for row_dict in records:
                                            rid = str(row_dict.get(id_col, ""))
                                            if rid and rid != "nan":
                                                row_dict["id"] = rid
                                                existing_respondents_by_id[rid] = row_dict
                                        print(f"[+] Loaded {len(existing_respondents_by_id)} existing respondent(s) from data.xlsx checkpoint.")
                                except Exception as e:
                                    print(f"[!] Warning reading existing Excel checkpoint: {e}")

                            json_dir = os.path.join(RESULT_DIR, str(survey_id), str(period_id), "json")
                            if os.path.exists(json_dir):
                                for j_file in os.listdir(json_dir):
                                    if j_file.endswith(".json"):
                                        j_id = j_file[:-5]
                                        if j_id not in existing_respondents_by_id:
                                            existing_respondents_by_id[j_id] = {"id": j_id}

                            period_respondents_map = dict(existing_respondents_by_id)

                            if not all_regions_completed:
                                for r_idx, path in enumerate(region_paths):
                                    region_key = " > ".join([r.get("fullCode", r.get("id", "")) for r in path])
                                    region_name_desc = " > ".join([r.get("name", "") for r in path])

                                    # Check if region is already recorded in region_progress.json
                                    if region_key in completed_region_keys:
                                        print(f"[*] Region [{region_name_desc}] is present in region_progress.json. Skipping server fetch.")
                                        continue

                                    if r_idx > 0:
                                        print(f"[*] Pausing {DEFAULT_DELAY}s before switching to next region...")
                                        time.sleep(DEFAULT_DELAY)

                                    # 1. Fetch all iteration of survey respondent for this specific region
                                    respondents = fetch_survey_respondents(driver, period_id, survey_id, region_path=path, delay=DEFAULT_DELAY)
                                    for item in respondents:
                                        item_id = str(item.get("id", ""))
                                        if item_id:
                                            period_respondents_map[item_id] = item

                                    # 2. Write/save respondent list to data.xlsx first
                                    print(f"[*] Writing respondents for [{region_name_desc}] to data.xlsx...")
                                    save_respondents_to_excel(list(period_respondents_map.values()), survey_id, period_id)

                                    # 3. Only after data.xlsx save succeeds, mark region as complete in region_progress.json
                                    completed_region_keys.add(region_key)
                                    progress["completed_regions"] = list(completed_region_keys)
                                    save_region_progress(survey_id, period_id, progress)
                                    print(f"[+] Region [{region_name_desc}] completed and saved to region_progress.json.")

                                if all_target_keys.issubset(completed_region_keys):
                                    print(f"\n[+] All regions for Period '{period_name}' completed.")
                            else:
                                print(f"[*] All regions for Period '{period_name}' were completed. Skipping region network requests entirely.")

                            period_respondents = list(period_respondents_map.values())

                            if period_respondents:
                                print(f"[*] Processing detail JSONs for {len(period_respondents)} respondent(s)...")
                                enrich_respondents_with_details(driver, period_respondents, survey_id, period_id, delay=DETAIL_DELAY)
                                save_respondents_to_excel(period_respondents, survey_id, period_id)

                            # Step 8: Extract dynamic respondent details to Excel
                            extract_json_to_excel(survey_id, period_id)

        # Keep browser open for inspection if needed
        input("\nPress Enter to close browser...")

    except Exception as e:
        print(f"[!] Error occurred: {e}")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()

