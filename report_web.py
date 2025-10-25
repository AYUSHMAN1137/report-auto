import os
import sys
import re
import time
import threading
import webbrowser
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional, List

from PyPDF2 import PdfReader, PdfWriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# Import dashboard (same module as provided)
from dashboard import ReportDashboard


# ===================== Configuration =====================
APP_ROOT = Path(__file__).resolve().parent
BASE_URL = "https://thyronxt.thyrocare.com/"
LOGIN_URL = "https://thyronxt.thyrocare.com/login"

# Credentials via env with safe defaults (change via env in production)
USERNAME = os.environ.get("THYROCARE_USERNAME", "ad051")
PASSWORD = os.environ.get("THYROCARE_PASSWORD", "1CNK3@wf1")

# WhatsApp Configuration
WHATSAPP_CONTACT = os.environ.get("WHATSAPP_CONTACT", "Ayushman1137")

# Folders
DOWNLOAD_DIR = APP_ROOT / "download"
TEMP_DIR = DOWNLOAD_DIR / "temp"
FINAL_DIR = DOWNLOAD_DIR / "final"
PROFILE_DIR = APP_ROOT / "profile_min"  # Minimal profile for login-only storage

for p in [DOWNLOAD_DIR, TEMP_DIR, FINAL_DIR, PROFILE_DIR]:
    p.mkdir(parents=True, exist_ok=True)

# Screenshot folder
SCREENSHOT_DIR = DOWNLOAD_DIR / "screenshots"
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

# Detect platform modifier key
IS_MAC = sys.platform == "darwin"
SELECT_ALL_MOD = Keys.COMMAND if IS_MAC else Keys.CONTROL

# Timeouts
TIMEOUTS = {
    "PAGE_LOAD": 60,
    "WAIT_SHORT": 10,
    "WAIT_MED": 20,
    "WAIT_LONG": 60,
    "DL_NEW_FILE": 120,
    "DL_ALL": 180,
}


# ===================== Profile Cleaner =====================
def clean_profile_bloat(dashboard=None):
    """
    Auto-clean cache/bloat folders from profile to keep size minimal.
    Preserves only: Cookies, Local Storage, IndexedDB (login data).
    Returns: dict with cleanup stats
    """
    if not PROFILE_DIR.exists():
        return {"cleaned": 0, "size_freed_mb": 0}
    
    # Calculate size before cleanup
    try:
        size_before = sum(f.stat().st_size for f in PROFILE_DIR.rglob('*') if f.is_file())
    except Exception:
        size_before = 0
    
    cleaned_count = 0
    
    # Top-level bloat folders to delete
    bloat_dirs = [
        "GrShaderCache", "GraphiteDawnCache", "ShaderCache", "GPUCache",
        "Crashpad", "Snapshots", "component_crx_cache", "extensions_crx_cache",
        "optimization_guide_model_store", "BrowserMetrics", "DeferredBrowserMetrics",
        "CertificateRevocation", "Crowd Deny", "WasmTtsEngine", "WidevineCdm",
        "MediaFoundationWidevineCdm", "OnDeviceHeadSuggestModel", "hyphen-data",
        "segmentation_platform", "AutofillStates", "CookieReadinessList",
        "FileTypePolicies", "FirstPartySetsPreloaded", "MEIPreload",
        "OpenCookieDatabase", "OptimizationHints", "OriginTrials", "PKIMetadata",
        "PrivacySandboxAttestationsPreloaded", "ProbabilisticRevealTokenRegistry",
        "RecoveryImproved", "SSLErrorAssistant", "Safe Browsing", "SafetyTips",
        "Subresource Filter", "TpcdMetadata", "TrustTokenKeyCommitments", "ZxcvbnData"
    ]
    
    for dirname in bloat_dirs:
        try:
            bloat_path = PROFILE_DIR / dirname
            if bloat_path.exists():
                shutil.rmtree(bloat_path, ignore_errors=True)
                cleaned_count += 1
        except Exception:
            pass
    
    # Top-level bloat files
    bloat_files = [
        "chrome_debug.log", "Breadcrumbs", "CrashpadMetrics-active.pma",
        "DevToolsActivePort", "First Run", "Last Browser", "Last Version",
        "Variations", "first_party_sets.db", "first_party_sets.db-journal"
    ]
    
    for filename in bloat_files:
        try:
            bloat_file = PROFILE_DIR / filename
            if bloat_file.exists():
                bloat_file.unlink()
                cleaned_count += 1
        except Exception:
            pass
    
    # Default profile cache folders (preserve Cookies, Local Storage, IndexedDB)
    default_dir = PROFILE_DIR / "Default"
    if default_dir.exists():
        default_bloat = [
            "Cache", "Code Cache", "GPUCache", "Service Worker",
            "DawnCache", "optimization_guide_hint_cache_store",
            "optimization_guide_prediction_model_downloads",
            "shared_proto_db", "VideoDecodeStats", "WebStorage",
            "blob_storage", "File System", "Platform Notifications",
            "Session Storage", "Sync Data", "Sync Extension Settings",
            "Extension State", "Extension Rules", "Extensions",
            "BudgetDatabase", "databases", "GCM Store", "QuotaManager",
            "QuotaManager-journal", "TransportSecurity", "Reporting and NEL",
            "Network Persistent State", "Affiliation Database",
            "Favicons", "Favicons-journal", "History", "History-journal",
            "Top Sites", "Top Sites-journal", "Visited Links"
        ]
        
        for dirname in default_bloat:
            try:
                bloat_path = default_dir / dirname
                if bloat_path.exists():
                    if bloat_path.is_dir():
                        shutil.rmtree(bloat_path, ignore_errors=True)
                    else:
                        bloat_path.unlink()
                    cleaned_count += 1
            except Exception:
                pass
    
    # Calculate size after cleanup
    try:
        size_after = sum(f.stat().st_size for f in PROFILE_DIR.rglob('*') if f.is_file())
        size_freed_mb = (size_before - size_after) / (1024 * 1024)
    except Exception:
        size_freed_mb = 0
    
    result = {"cleaned": cleaned_count, "size_freed_mb": round(size_freed_mb, 2)}
    
    if dashboard and cleaned_count > 0:
        dashboard.add_log(f"🧹 Profile cleaned: {cleaned_count} items removed, {result['size_freed_mb']:.2f} MB freed")
    
    return result


# ===================== Helpers =====================
def take_error_screenshot(driver, error_name: str) -> Optional[Path]:
    try:
        if not driver:
            return None
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"error_{sanitize_filename(error_name)}_{timestamp}.png"
        screenshot_path = SCREENSHOT_DIR / filename
        driver.save_screenshot(str(screenshot_path))
        print(f" Screenshot saved: {screenshot_path}")
        return screenshot_path
    except Exception as e:
        print(f" Screenshot failed: {e}")
        return None


def sanitize_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = re.sub(r"\s+", " ", name)
    return (name or "file")[:150]


def unique_path(dir_path: Path, base: str, ext: str) -> Path:
    p = dir_path / f"{base}{ext}"
    if not p.exists():
        return p
    n = 1
    while True:
        p2 = dir_path / f"{base} ({n}){ext}"
        if not p2.exists():
            return p2
        n += 1


def _is_file_stable(path: Path, interval: float = 0.4) -> bool:
    try:
        s1 = path.stat().st_size
        time.sleep(interval)
        s2 = path.stat().st_size
        return s1 == s2
    except Exception:
        return False


def wait_for_new_download(
    dir_path: Path,
    start_ts: float,
    timeout: int = TIMEOUTS["DL_NEW_FILE"],
    cancel: Optional[threading.Event] = None,
) -> Optional[Path]:
    end = time.time() + timeout
    last_candidate: Optional[Path] = None
    last_size = -1
    stable_ticks = 0

    while time.time() < end:
        if cancel and cancel.is_set():
            return None

        if any(f.suffix == ".crdownload" for f in dir_path.glob("*.crdownload")):
            time.sleep(0.25)
            continue

        files = [f for f in dir_path.glob("*") if f.is_file() and f.suffix != ".crdownload"]
        candidates = [f for f in files if f.stat().st_mtime >= start_ts - 1]

        if candidates:
            newest = max(candidates, key=lambda f: f.stat().st_mtime)
            try:
                size = newest.stat().st_size
            except Exception:
                time.sleep(0.25)
                continue

            if last_candidate and newest == last_candidate and size == last_size:
                stable_ticks += 1
                if stable_ticks >= 3:
                    return newest
            else:
                last_candidate = newest
                last_size = size
                stable_ticks = 0

        time.sleep(0.25)

    return None


def wait_all_downloads_complete(
    dir_path: Path,
    timeout: int = TIMEOUTS["DL_ALL"],
    cancel: Optional[threading.Event] = None,
) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        if cancel and cancel.is_set():
            return False
        if not any(f.suffix == ".crdownload" for f in dir_path.glob("*.crdownload")):
            return True
        time.sleep(0.25)
    return False


# ===================== Selenium: Driver / Login =====================
def create_driver(dashboard=None):
    # Clean profile bloat before starting browser
    if dashboard:
        dashboard.add_log("🧹 Cleaning profile cache...")
    clean_profile_bloat(dashboard)
    
    opts = webdriver.ChromeOptions()
    # Profile so WhatsApp stays logged in
    opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
    opts.add_argument("--profile-directory=Default")

    # Less noisy / automation flags
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--start-maximized")
    opts.add_argument("--log-level=3")
    opts.add_argument("--disable-logging")
    opts.add_argument("--silent")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-sync")
    opts.add_argument("--metrics-recording-only")
    opts.add_argument("--disable-default-apps")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--lang=en-US,en;q=0.9")
    
    # ===== CACHE LIMITING FLAGS (MINIMAL STORAGE) =====
    # Limit disk cache to ~10MB
    opts.add_argument("--disk-cache-size=10485760")
    # Disable media cache
    opts.add_argument("--media-cache-size=1")
    # Disable application cache
    opts.add_argument("--disable-application-cache")
    # Aggressive cache discard
    opts.add_argument("--aggressive-cache-discard")
    # Disable component updates (prevents component_crx_cache bloat)
    opts.add_argument("--disable-component-update")
    # Disable optimization hints and privacy sandbox features
    opts.add_argument("--disable-features=OptimizationHints,OptimizationGuideModelDownloading,OptimizationHintsFetching,InterestCohortFeature,PrivacySandboxSettings4,FledgePst,Prerender2")
    # Disable background services that create cache
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--disable-renderer-backgrounding")
    # Disable crash reporting (prevents Crashpad bloat)
    opts.add_argument("--disable-breakpad")
    # Disable client side phishing detection (reduces Safe Browsing cache)
    opts.add_argument("--disable-client-side-phishing-detection")
    # Disable domain reliability (reduces telemetry)
    opts.add_argument("--disable-domain-reliability")

    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    opts.add_experimental_option("prefs", prefs)

    driver = None
    try:
        # Selenium 4.6+ will auto-manage the driver via Selenium Manager
        service = Service(log_output=os.devnull)
        driver = webdriver.Chrome(service=service, options=opts)
    except Exception as e:
        # Fallback to local chromedriver next to script
        exe_name = "chromedriver.exe" if os.name == "nt" else "chromedriver"
        local_driver = APP_ROOT / exe_name
        if local_driver.exists():
            service = Service(str(local_driver), log_output=os.devnull)
            driver = webdriver.Chrome(service=service, options=opts)
        else:
            raise RuntimeError(
                "ChromeDriver not found. Place chromedriver(.exe) next to this script "
                "or use Selenium 4.6+ so it can auto-manage the driver."
            ) from e

    driver.set_page_load_timeout(TIMEOUTS["PAGE_LOAD"])

    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(DOWNLOAD_DIR)},
        )
    except Exception:
        pass

    wait = WebDriverWait(driver, TIMEOUTS["WAIT_MED"])
    return driver, wait


def is_login_page(driver) -> bool:
    try:
        if "/login" in driver.current_url:
            return True
        u = driver.find_elements(
            By.XPATH,
            "//input[@placeholder='Enter Username' or @name='username' or @id='username']",
        )
        p = driver.find_elements(
            By.XPATH,
            "//input[@placeholder='Enter Password' or @type='password' or @name='password' or @id='password']",
        )
        return any(e.is_displayed() for e in u) and any(e.is_displayed() for e in p)
    except Exception:
        return False


def perform_login(driver, wait):
    driver.get(LOGIN_URL)
    user = wait.until(
        EC.visibility_of_element_located(
            (
                By.XPATH,
                "//input[@placeholder='Enter Username' or @name='username' or @id='username']",
            )
        )
    )
    pwd = wait.until(
        EC.visibility_of_element_located(
            (
                By.XPATH,
                "//input[@placeholder='Enter Password' or @type='password' or @name='password' or @id='password']",
            )
        )
    )

    user.clear()
    user.send_keys(USERNAME)
    pwd.clear()
    pwd.send_keys(PASSWORD)

    clicked = False
    for xp in [
        "//button[@id='loginBtn']",
        "//button[@type='submit']",
        "//button[contains(., 'SIGN IN') or contains(., 'Sign In')]",
    ]:
        try:
            btn = driver.find_element(By.XPATH, xp)
            if btn.is_enabled() and btn.is_displayed():
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp)))
                    btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                clicked = True
                break
        except NoSuchElementException:
            continue

    if not clicked:
        pwd.send_keys(Keys.ENTER)

    try:
        WebDriverWait(driver, 15).until(lambda d: "/login" not in d.current_url)
    except TimeoutException:
        pass


def ensure_logged_in(driver, wait) -> bool:
    if "thyronxt.thyrocare.com" not in driver.current_url:
        driver.get(BASE_URL)
        time.sleep(1)
    if is_login_page(driver):
        perform_login(driver, wait)
        time.sleep(1)
    if is_login_page(driver):
        perform_login(driver, wait)
        time.sleep(1)
    return not is_login_page(driver)


def maybe_close_post_login_popup(driver, delay_after_login=2):
    time.sleep(delay_after_login)
    close_xpaths = [
        "//*[@aria-label='Close']",
        "//button[@title='Close']",
        "//button[contains(@class,'btn-close') or contains(@class,'close')]",
        "//div[contains(@class,'modal')]//button[contains(@class,'close') or contains(@class,'btn-close')]",
        "//button[normalize-space()='×' or normalize-space()='X' or normalize-space()='x']",
        "//span[normalize-space()='×' or normalize-space()='X']/parent::button",
    ]
    for xp in close_xpaths:
        try:
            for b in driver.find_elements(By.XPATH, xp):
                if b.is_displayed():
                    try:
                        b.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", b)
                    time.sleep(0.2)
                    return
        except Exception:
            continue


def _click_first_visible(driver, xpaths: List[str]) -> bool:
    for xp in xpaths:
        try:
            for el in driver.find_elements(By.XPATH, xp):
                if not el.is_displayed():
                    continue
                target = el
                if el.tag_name.lower() not in ("a", "button"):
                    try:
                        target = el.find_element(
                            By.XPATH, "./ancestor::*[self::a or self::button][1]"
                        )
                    except Exception:
                        pass
                try:
                    target.click()
                except Exception:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
                    time.sleep(0.1)
                    driver.execute_script("arguments[0].click();", target)
                time.sleep(0.2)
                return True
        except Exception:
            continue
    return False


def _wait_for_orders_ready(driver, wait) -> bool:
    try:
        wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//input[@placeholder='Enter Barcode' or "
                    "contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'barcode')]",
                )
            )
        )
        return True
    except TimeoutException:
        return "order" in driver.current_url.lower()


def go_to_orders(driver, wait) -> bool:
    orders_xpaths = [
        "//a[normalize-space()='Orders' or normalize-space()='Orders & Leads']",
        "//button[normalize-space()='Orders']",
        "//span[normalize-space()='Orders']",
        "//a[.//span[normalize-space()='Orders']]",
        "//a[contains(translate(@href,'ORDERS','orders'),'orders')]",
    ]
    if _click_first_visible(driver, orders_xpaths):
        return _wait_for_orders_ready(driver, wait)

    # Try menu first
    for xp in [
        "//*[@aria-label='Menu' or @aria-label='Open Menu']",
        "//button[contains(@class,'menu') or contains(@class,'toggle')]",
    ]:
        try:
            for el in driver.find_elements(By.XPATH, xp):
                if el.is_displayed():
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                    time.sleep(0.3)
                    if _click_first_visible(driver, orders_xpaths):
                        return _wait_for_orders_ready(driver, wait)
        except Exception:
            continue

    try:
        driver.get(BASE_URL + "orders")
        return _wait_for_orders_ready(driver, wait)
    except Exception:
        pass
    return False


# ===================== Orders Page Actions =====================
def find_barcode_input(driver, wait):
    xps = [
        "//input[@placeholder='Enter Barcode']",
        "//div[.//span[contains(normalize-space(),'Barcode')] or .//label[contains(normalize-space(),'Barcode')]]//input",
        "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'barcode')]",
        "//input[@name='barcode' or @id='barcode']",
    ]
    for xp in xps:
        try:
            el = wait.until(EC.visibility_of_element_located((By.XPATH, xp)))
            if el.is_displayed():
                return el
        except TimeoutException:
            continue
    return None


def click_go_next_to_input(driver, input_el) -> bool:
    patterns = [
        "./following-sibling::*//button[normalize-space()='GO'][1]",
        "./parent::*//button[normalize-space()='GO'][1]",
        "./ancestor::*[self::div or self::form][contains(@class,'input') or contains(@class,'filter') or contains(@class,'toolbar')][1]//button[normalize-space()='GO']",
        "//div[.//input[@placeholder='Enter Barcode' or contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'barcode')]]//button[normalize-space()='GO']",
    ]
    for rel in patterns:
        try:
            btn = (
                input_el.find_element(By.XPATH, rel)
                if rel.startswith("./")
                else driver.find_element(By.XPATH, rel)
            )
            if btn.is_displayed():
                try:
                    btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                return True
        except Exception:
            continue

    try:
        btn = driver.find_element(By.XPATH, "//button[normalize-space()='GO']")
        if btn.is_displayed():
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
            return True
    except Exception:
        pass
    return False


def wait_for_result_rows(driver, timeout=12, cancel: Optional[threading.Event] = None):
    end = time.time() + timeout
    while time.time() < end:
        if cancel and cancel.is_set():
            return []
        try:
            rows = [r for r in driver.find_elements(By.XPATH, "//table//tbody/tr") if r.is_displayed()]
            if rows:
                return rows
            empty = driver.find_elements(
                By.XPATH,
                "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'no record') or "
                "contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'no results') or "
                "contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'no data') or "
                "contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'not found') or "
                "contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'invalid barcode')]",
            )
            if any(e.is_displayed() for e in empty):
                return []
        except Exception:
            pass
        time.sleep(0.3)
    return []


def extract_patient_name_from_row(row) -> Optional[str]:
    # Try common clickable/name spots
    try:
        el = row.find_element(
            By.XPATH,
            ".//*[self::div or self::a or self::span][contains(@class,'clickable')][normalize-space()]",
        )
        txt = el.text.strip()
        if txt:
            return txt
    except Exception:
        pass
    try:
        el = row.find_element(By.XPATH, ".//a[normalize-space() and not(.//*[name()='svg'])]")
        txt = el.text.strip()
        if txt:
            return txt
    except Exception:
        pass
    try:
        tds = row.find_elements(By.TAG_NAME, "td")
        for td in tds:
            text = td.text.strip()
            if text:
                first = text.splitlines()[0].strip()
                if first:
                    return first
    except Exception:
        pass
    return None


def click_row_download_icon(driver, row, timeout=15) -> bool:
    """Enhanced download icon detection with better retries and more patterns."""
    actions = ActionChains(driver)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
        time.sleep(0.3)
        actions.move_to_element(row).perform()
        time.sleep(0.2)
    except Exception:
        pass

    end = time.time() + timeout
    attempt = 0
    while time.time() < end:
        attempt += 1
        
        # Method 1: data-tip attribute
        try:
            btns = row.find_elements(By.CSS_SELECTOR, "button[data-tip^='download_']")
            for b in btns:
                if b.is_displayed():
                    try:
                        b.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", b)
                    return True
        except Exception:
            pass

        # Method 2: Multiple XPath patterns for download icons
        xps = [
            ".//*[self::a or self::button][.//*[name()='svg' and @data-icon='download']]",
            ".//td[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'action')]//*[self::a or self::button][.//*[name()='svg' and @data-icon='download']]",
            ".//*[self::a or self::button][.//*[name()='svg'][contains(@class, 'download')]]",
            ".//*[contains(@class, 'download')][self::a or self::button]",
            ".//button[@title='Download' or @aria-label='Download']",
            ".//a[@title='Download' or @aria-label='Download']",
            ".//*[name()='svg'][@data-icon='download']/ancestor::button[1]",
            ".//*[name()='svg'][@data-icon='download']/ancestor::a[1]",
        ]
        for xp in xps:
            try:
                el = row.find_element(By.XPATH, xp)
                if el.is_displayed():
                    try:
                        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, xp)))
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        time.sleep(0.15)
                        driver.execute_script("arguments[0].click();", el)
                    return True
            except Exception:
                continue

        # Method 3: Try hovering again every few attempts
        if attempt % 3 == 0:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
                time.sleep(0.2)
                actions.move_to_element(row).perform()
                time.sleep(0.3)
            except Exception:
                pass
        
        time.sleep(0.3)
    return False


# ===================== PDF Post-processing =====================
def process_single_pdf(src_pdf: Path) -> bool:
    """Process PDF by removing first 2 and last page. Returns True if processed successfully."""
    try:
        # Ensure TEMP_DIR exists
        TEMP_DIR.mkdir(parents=True, exist_ok=True)
        
        if src_pdf.suffix.lower() != ".pdf":
            print(f" {src_pdf.name}: Not a PDF file")
            return False

        # Read and decrypt if needed
        reader = PdfReader(str(src_pdf))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception as e:
                print(f" {src_pdf.name}: Decryption failed - {e}")
                # Try to copy as-is if decryption fails
                try:
                    shutil.copy2(src_pdf, TEMP_DIR / src_pdf.name)
                    src_pdf.unlink()
                    return True
                except Exception:
                    return False

        total = len(reader.pages)
        
        # Handle different PDF sizes
        if total <= 3:
            # For small PDFs (1-3 pages), copy as-is
            try:
                shutil.copy2(src_pdf, TEMP_DIR / src_pdf.name)
                src_pdf.unlink()
                print(f" {src_pdf.name}: Too small ({total} pages), copied as-is")
                return True
            except Exception as e:
                print(f" {src_pdf.name}: Copy failed - {e}")
                return False
        
        # For PDFs with >3 pages, remove first 2 and last page
        writer = PdfWriter()
        for i in range(2, total - 1):
            writer.add_page(reader.pages[i])

        out_path = TEMP_DIR / src_pdf.name
        with out_path.open("wb") as f:
            writer.write(f)

        try:
            src_pdf.unlink()
        except Exception:
            pass

        print(f" {src_pdf.name}: Processed ({total} -> {total-3} pages)")
        return True
    except Exception as e:
        print(f" {src_pdf.name}: Processing error - {e}")
        # Fallback: try to copy as-is
        try:
            TEMP_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src_pdf, TEMP_DIR / src_pdf.name)
            src_pdf.unlink()
            print(f" {src_pdf.name}: Copied as-is after error")
            return True
        except Exception:
            return False


def merge_temp_pdfs(output_name: Optional[str] = None) -> Optional[Path]:
    pdfs = sorted(TEMP_DIR.glob("*.pdf"), key=lambda p: p.name.lower())
    if not pdfs:
        return None

    if not output_name:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"merged_reports_{ts}.pdf"

    output_path = FINAL_DIR / output_name
    writer = PdfWriter()
    added_pages = 0

    for p in pdfs:
        try:
            # Open with context manager to avoid file-handle leaks
            with p.open("rb") as fh:
                r = PdfReader(fh)
                if r.is_encrypted:
                    try:
                        r.decrypt("")
                    except Exception:
                        continue
                for page in r.pages:
                    writer.add_page(page)
                    added_pages += 1
        except Exception:
            continue

    if added_pages == 0:
        return None

    with output_path.open("wb") as f:
        writer.write(f)

    return output_path


def clean_temp():
    for f in TEMP_DIR.glob("*.pdf"):
        try:
            f.unlink()
        except Exception:
            pass


# ===================== WhatsApp Automation =====================
def send_pdf_via_whatsapp(
    driver,
    wait,
    pdf_path: Path,
    contact_name: str,
    dashboard,
    cancel_ev: Optional[threading.Event] = None,
) -> bool:
    try:
        if cancel_ev and cancel_ev.is_set():
            return False

        dashboard.add_log("📱 Opening WhatsApp Web...")
        dashboard.update_stats(
            {"current_step": "WhatsApp • Opening", "current_step_index": 1, "current_step_total": 5}
        )

        driver.get("https://web.whatsapp.com")

        dashboard.add_log("⏳ Waiting for WhatsApp to load...")
        dashboard.update_stats({"current_step": "WhatsApp • Loading", "current_step_index": 2})

        try:
            WebDriverWait(driver, 60).until(
                lambda d: d.find_elements(By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
                or d.find_elements(By.XPATH, "//canvas[@aria-label='Scan me!']")
            )

            qr_codes = driver.find_elements(By.XPATH, "//canvas[@aria-label='Scan me!']")
            if qr_codes and any(q.is_displayed() for q in qr_codes):
                dashboard.add_log("📷 QR Code detected! Please scan...")
                WebDriverWait(driver, 120).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
                    )
                )
                dashboard.add_log("✅ QR Code scanned!")
                time.sleep(2)

        except TimeoutException:
            dashboard.add_log("❌ WhatsApp failed to load")
            return False

        if cancel_ev and cancel_ev.is_set():
            return False

        dashboard.add_log(f"🔍 Searching: {contact_name}")
        dashboard.update_stats({"current_step": "WhatsApp • Searching", "current_step_index": 3})

        try:
            search_box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
            search_box.click()
            time.sleep(0.3)

            # Use platform-specific select-all
            search_box.send_keys(SELECT_ALL_MOD, "a")
            search_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)
            search_box.send_keys(contact_name)
            time.sleep(1.8)

            contact = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, f"//span[@title='{contact_name}']"))
            )
            contact.click()
            time.sleep(1.2)

            dashboard.add_log(f"✅ Chat opened: {contact_name}")

        except TimeoutException:
            dashboard.add_log(f"❌ Contact not found: {contact_name}")
            return False

        if cancel_ev and cancel_ev.is_set():
            return False

        dashboard.add_log(f"📎 Attaching: {pdf_path.name}")
        dashboard.update_stats({"current_step": "WhatsApp • Attaching", "current_step_index": 4})

        try:
            attach_xpaths = [
                "//div[@title='Attach']",
                "//div[@aria-label='Attach']",
                "//span[@data-icon='plus']",
                "//span[@data-icon='attach-menu-plus']",
            ]

            attached = False
            for xpath in attach_xpaths:
                try:
                    attach_btn = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    if attach_btn and attach_btn.is_displayed():
                        attach_btn.click()
                        attached = True
                        break
                except Exception:
                    continue

            if not attached:
                dashboard.add_log("❌ Attach button not found")
                return False

            time.sleep(0.6)

            file_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@accept='*' and @type='file']"))
            )
            file_input.send_keys(str(pdf_path.absolute()))
            dashboard.add_log("⏳ Uploading...")
            time.sleep(3)
            dashboard.add_log("✅ File uploaded")

        except Exception as e:
            dashboard.add_log(f"❌ Attach failed: {e}")
            return False

        if cancel_ev and cancel_ev.is_set():
            return False

        dashboard.add_log("📤 Sending PDF...")
        dashboard.update_stats({"current_step": "WhatsApp • Sending", "current_step_index": 5})

        try:
            dashboard.add_log("⏳ Waiting for document preview...")
            time.sleep(2.5)

            send_button_xpaths = [
                "//div[@role='button'][@aria-label='Send']",
                "//button[@aria-label='Send']",
                "//span[@data-icon='send'][@data-testid='send']",
                "//span[@data-icon='send']/parent::button",
                "//span[@data-icon='send']/parent::div[@role='button']",
                "//div[contains(@class, 'copyable-area')]//span[@data-icon='send']",
                "//footer//span[@data-icon='send']",
                "//*[@data-testid='send']",
                "//button[.//span[@data-icon='send']]",
            ]

            send_success = False
            actions = ActionChains(driver)

            for xpath in send_button_xpaths:
                try:
                    send_btn = WebDriverWait(driver, 4).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    if send_btn and send_btn.is_displayed():
                        driver.execute_script(
                            "arguments[0].scrollIntoView({block:'center', behavior:'smooth'});", send_btn
                        )
                        time.sleep(0.3)
                        try:
                            WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.XPATH, xpath))
                            ).click()
                        except Exception:
                            try:
                                actions.move_to_element(send_btn).click().perform()
                            except Exception:
                                driver.execute_script("arguments[0].click();", send_btn)
                        send_success = True
                        break
                except Exception:
                    continue

            if not send_success:
                dashboard.add_log("🔄 Trying ENTER key...")
                caption_selectors = [
                    "//div[@contenteditable='true'][@data-tab='10']",
                    "//div[@contenteditable='true' and contains(@class, 'copyable-text')]",
                    "//div[@role='textbox'][@contenteditable='true']",
                ]
                for selector in caption_selectors:
                    try:
                        caption = driver.find_element(By.XPATH, selector)
                        if caption.is_displayed():
                            caption.click()
                            time.sleep(0.2)
                            caption.send_keys(Keys.ENTER)
                            send_success = True
                            break
                    except Exception:
                        continue

            if send_success:
                time.sleep(2.5)
                dashboard.add_log(f"🎉 PDF sent to {contact_name}!")
                return True
            else:
                dashboard.add_log("❌ Send button not clickable")
                return False

        except Exception as e:
            dashboard.add_log(f"❌ Send error: {e}")
            return False

    except Exception as e:
        dashboard.add_log(f"❌ WhatsApp automation error: {e}")
        return False


# ===================== Main Application =====================
class ReportAutomation:
    def __init__(self):
        self.dashboard = ReportDashboard(port=5001)
        self.dashboard.set_download_folder(DOWNLOAD_DIR)
        self.dashboard.set_whatsapp_contact(WHATSAPP_CONTACT)
        self.dashboard.set_handlers(on_control=self.handle_control)

        self.driver = None
        self.wait = None
        self.cancel_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self.is_running = False

    def safe_quit_driver(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        finally:
            self.driver = None

    def handle_control(self, action: str, data: dict):
        if action == "start":
            barcodes_input = data.get("barcodes")
            
            # Handle both list (from frontend) and string (for compatibility)
            if isinstance(barcodes_input, list):
                # Frontend sends array - use it directly
                barcodes = [str(b).strip() for b in barcodes_input if str(b).strip()]
            elif isinstance(barcodes_input, str):
                # String format - parse it
                barcodes_text = barcodes_input.strip()
                if not barcodes_text:
                    return {"ok": False, "error": "No barcodes provided"}
                parts = re.split(r"[,;\n\r]+", barcodes_text)
                barcodes = [b.strip() for b in parts if b.strip()]
            else:
                return {"ok": False, "error": "No barcodes provided"}

            if not barcodes:
                return {"ok": False, "error": "No valid barcodes"}

            # Remove duplicates, preserve order
            seen, unique_barcodes = set(), []
            for b in barcodes:
                if b not in seen:
                    seen.add(b)
                    unique_barcodes.append(b)

            if not unique_barcodes:
                return {"ok": False, "error": "No valid barcodes"}

            if self.is_running:
                return {"ok": False, "error": "Already running"}

            # Start worker
            self.cancel_event.clear()
            self.worker_thread = threading.Thread(
                target=self.worker, args=(unique_barcodes,), daemon=True
            )
            self.worker_thread.start()
            return {"ok": True}

        elif action == "cancel":
            if self.is_running:
                self.cancel_event.set()
                self.dashboard.add_log("⏹ Cancelling...")
                self.safe_quit_driver()
                return {"ok": True}
            return {"ok": False, "error": "Not running"}
        
        elif action == "clean_profile":
            if self.is_running:
                return {"ok": False, "error": "Cannot clean profile while automation is running"}
            
            try:
                self.dashboard.add_log("🧹 Manual profile cleanup initiated...")
                result = clean_profile_bloat(self.dashboard)
                if result["cleaned"] == 0:
                    self.dashboard.add_log("✅ Profile already clean (no bloat found)")
                return {"ok": True, "result": result}
            except Exception as e:
                error_msg = f"Profile cleanup failed: {str(e)}"
                self.dashboard.add_log(f"❌ {error_msg}")
                return {"ok": False, "error": error_msg}

        return {"ok": False, "error": f"Unknown action: {action}"}

    def worker(self, barcodes: List[str]):
        self.is_running = True

        # Reset all stats
        self.dashboard.update_stats(
            {
                "status": "running",
                "total_barcodes": 0,
                "done_barcodes": 0,
                "total_pdfs": 0,
                "done_pdfs": 0,
                "current_step": "Initializing...",
                "current_step_index": 0,
                "current_step_total": 5,
                "current_barcode": "-",
                "whatsapp_status": "pending",
            }
        )

        self.dashboard.add_log(f"🚀 Starting with {len(barcodes)} barcode(s)...")

        # Start browser
        try:
            self.dashboard.add_log("🌐 Starting browser...")
            self.driver, self.wait = create_driver(self.dashboard)
            self.dashboard.add_log("✅ Browser started successfully")
        except Exception as e:
            error_msg = str(e)
            self.dashboard.add_log(f"❌ Browser error: {error_msg}")
            self.dashboard.send_error("Browser Startup Failed", error_msg)
            self._finish()
            return

        try:
            if self.cancel_event.is_set():
                self._cancelled()
                return

            # Login
            self.dashboard.add_log("🔐 Logging in to Thyronxt...")
            try:
                ok = ensure_logged_in(self.driver, self.wait)
                if ok:
                    self.dashboard.add_log("✅ Login successful!")
                else:
                    self.dashboard.add_log("❌ Login failed - check credentials")
                    screenshot = take_error_screenshot(self.driver, "login_failed")
                    self.dashboard.send_error(
                        "Login Failed",
                        "Unable to login. Check credentials or site availability.",
                        str(screenshot) if screenshot else None,
                    )
                    self._finish()
                    return
            except Exception as e:
                error_msg = str(e)
                self.dashboard.add_log(f"❌ Login error: {error_msg}")
                screenshot = take_error_screenshot(self.driver, "login_error")
                self.dashboard.send_error(
                    "Login Error", error_msg, str(screenshot) if screenshot else None
                )
                self._finish()
                return

            if self.cancel_event.is_set():
                self._cancelled()
                return

            # Navigate to Orders
            self.dashboard.add_log("🔄 Navigating to Orders page...")
            try:
                maybe_close_post_login_popup(self.driver, delay_after_login=2)
                orders_ok = go_to_orders(self.driver, self.wait)
                if orders_ok:
                    self.dashboard.add_log("✅ Orders page ready")
                else:
                    self.dashboard.add_log("⚠️ Orders page navigation issue - continuing anyway")
            except Exception as e:
                self.dashboard.add_log(f"⚠️ Orders navigation warning: {str(e)} - continuing...")

            # Process barcodes
            self.dashboard.update_stats({"total_barcodes": len(barcodes), "done_barcodes": 0})
            self.dashboard.add_log(f"📋 Processing {len(barcodes)} barcode(s)...")

            for idx, bc in enumerate(barcodes, start=1):
                if self.cancel_event.is_set():
                    self.dashboard.add_log("⏹ Cancelled")
                    break

                self.dashboard.update_stats({"current_barcode": bc})

                # 1: Orders
                self.dashboard.update_stats(
                    {"current_step": f"{bc} • Orders", "current_step_index": 1, "current_step_total": 5}
                )
                if not go_to_orders(self.driver, self.wait):
                    error_msg = f"{bc}: Orders page unavailable"
                    self.dashboard.add_log(f"❌ {error_msg}")
                    self.dashboard.send_error("Navigation Error", error_msg)
                    self.dashboard.update_stats({"done_barcodes": idx})
                    continue

                if self.cancel_event.is_set():
                    break

                # 2: Search
                self.dashboard.update_stats({"current_step": f"{bc} • Search", "current_step_index": 2})
                inp = find_barcode_input(self.driver, self.wait)
                if not inp:
                    error_msg = f"{bc}: Search input not found"
                    self.dashboard.add_log(f"❌ {error_msg}")
                    self.dashboard.send_error("Search Error", error_msg)
                    self.dashboard.update_stats({"done_barcodes": idx})
                    continue

                try:
                    inp.click()
                    inp.send_keys(SELECT_ALL_MOD, "a")
                    inp.send_keys(Keys.BACKSPACE)
                    time.sleep(0.1)
                    inp.send_keys(bc)
                    if not click_go_next_to_input(self.driver, inp):
                        inp.send_keys(Keys.ENTER)
                except Exception as e:
                    if self.cancel_event.is_set():
                        break
                    error_msg = f"{bc}: Search failed - {str(e)}"
                    self.dashboard.add_log(f"❌ {error_msg}")
                    self.dashboard.send_error("Search Error", error_msg)
                    self.dashboard.update_stats({"done_barcodes": idx})
                    continue

                if self.cancel_event.is_set():
                    break

                # 3: Wait for results
                self.dashboard.update_stats({"current_step": f"{bc} • Wait", "current_step_index": 3})
                rows = wait_for_result_rows(self.driver, timeout=12, cancel=self.cancel_event)
                if self.cancel_event.is_set():
                    break

                if not rows:
                    error_msg = f"{bc}: No results found (invalid barcode or no data)"
                    self.dashboard.add_log(f"⛔ {error_msg}")
                    self.dashboard.send_error("No Results", error_msg)
                    self.dashboard.update_stats({"done_barcodes": idx})
                    continue

                row = rows[0]
                patient_name = extract_patient_name_from_row(row) or bc
                patient_name = sanitize_filename(patient_name)

                # 4: Download
                self.dashboard.update_stats({"current_step": f"{bc} • Download", "current_step_index": 4})
                start_ts = time.time()
                if not click_row_download_icon(self.driver, row, timeout=15):
                    if self.cancel_event.is_set():
                        break
                    error_msg = f"{bc}: Download icon not found (may not be ready or report unavailable)"
                    self.dashboard.add_log(f"⚠️ {error_msg}")
                    screenshot = take_error_screenshot(self.driver, f"download_icon_missing_{bc}")
                    self.dashboard.send_error(
                        "Download Icon Missing", 
                        error_msg,
                        str(screenshot) if screenshot else None
                    )
                    self.dashboard.update_stats({"done_barcodes": idx})
                    continue

                # 5: Save
                self.dashboard.update_stats({"current_step": f"{bc} • Save", "current_step_index": 5})
                downloaded = wait_for_new_download(
                    DOWNLOAD_DIR, start_ts, timeout=TIMEOUTS["DL_NEW_FILE"], cancel=self.cancel_event
                )
                if self.cancel_event.is_set():
                    break

                if not downloaded:
                    error_msg = f"{bc}: Download timeout (file not received within {TIMEOUTS['DL_NEW_FILE']}s)"
                    self.dashboard.add_log(f"⚠️ {error_msg}")
                    screenshot = take_error_screenshot(self.driver, f"download_timeout_{bc}")
                    self.dashboard.send_error(
                        "Download Timeout",
                        error_msg,
                        str(screenshot) if screenshot else None
                    )
                    self.dashboard.update_stats({"done_barcodes": idx})
                    continue

                target = unique_path(DOWNLOAD_DIR, patient_name, downloaded.suffix)
                try:
                    downloaded.rename(target)
                except PermissionError:
                    time.sleep(1.0)
                    try:
                        downloaded.rename(target)
                    except Exception as e:
                        error_msg = f"{bc}: File rename failed - {str(e)}"
                        self.dashboard.add_log(f"⚠️ {error_msg}")
                        self.dashboard.send_error("File Error", error_msg)
                        self.dashboard.update_stats({"done_barcodes": idx})
                        continue

                self.dashboard.add_log(f"✅ {bc}: {target.name}")
                self.dashboard.update_stats({"done_barcodes": idx})

            # Wait for downloads to complete
            if not self.cancel_event.is_set():
                self.dashboard.add_log("⏳ Finalizing downloads...")
                wait_all_downloads_complete(DOWNLOAD_DIR, timeout=TIMEOUTS["DL_ALL"], cancel=self.cancel_event)

            if not self.cancel_event.is_set():
                self.dashboard.add_log("✅ Downloads complete")

        except Exception as e:
            if not self.cancel_event.is_set():
                error_msg = str(e)
                self.dashboard.add_log(f"❌ Error: {error_msg}")
                screenshot = take_error_screenshot(self.driver, "processing_error")
                self.dashboard.send_error(
                    "Processing Error", error_msg, str(screenshot) if screenshot else None
                )

        if self.cancel_event.is_set():
            self._cancelled()
            return

        # PDF Processing
        pdfs = [p for p in DOWNLOAD_DIR.glob("*.pdf") if p.is_file()]
        if not pdfs:
            self.dashboard.add_log("ℹ️ No PDFs to process")
            self._finish()
            return

        self.dashboard.add_log(f"📄 Processing {len(pdfs)} PDFs...")
        self.dashboard.update_stats({"total_pdfs": len(pdfs), "done_pdfs": 0})

        processed = 0
        skipped = 0
        for idx, p in enumerate(pdfs, start=1):
            if self.cancel_event.is_set():
                self._cancelled()
                return
            self.dashboard.update_stats({"current_step": f"PDF • {p.name}", "current_step_index": 0})
            if process_single_pdf(p):
                processed += 1
            else:
                # Fallback: if processing fails, try to copy as-is
                try:
                    TEMP_DIR.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(p, TEMP_DIR / p.name)
                    p.unlink()
                    processed += 1
                    self.dashboard.add_log(f"⚠️ {p.name}: Copied without processing")
                except Exception as e:
                    self.dashboard.add_log(f"❌ {p.name}: Failed - {e}")
                    skipped += 1
                    try:
                        p.unlink()
                    except Exception:
                        pass
            self.dashboard.update_stats({"done_pdfs": idx})

        if skipped > 0:
            self.dashboard.add_log(f"✅ Processed: {processed} PDFs ({skipped} skipped)")
        else:
            self.dashboard.add_log(f"✅ Processed: {processed} PDFs")

        if self.cancel_event.is_set():
            self._cancelled()
            return

        # Merge PDFs
        if processed > 0:
            self.dashboard.add_log("🧩 Merging PDFs...")
            merged = merge_temp_pdfs()

            if merged:
                clean_temp()
                self.dashboard.add_log(f"✅ Merged: {merged.name}")
            else:
                self.dashboard.add_log("❌ Merge failed - no valid PDFs in temp folder")
                self._finish()
                return
        else:
            self.dashboard.add_log("⚠️ No PDFs were processed successfully")
            self._finish()
            return

        if merged:

            # WhatsApp Integration
            if not self.cancel_event.is_set():
                self.dashboard.add_log("=" * 50)
                self.dashboard.add_log("🚀 WhatsApp automation starting...")
                self.dashboard.add_log("=" * 50)
                self.dashboard.update_stats({"status": "whatsapp", "whatsapp_status": "sending"})

                # Get the selected WhatsApp contact from dashboard
                whatsapp_contact = self.dashboard.get_whatsapp_contact()
                self.dashboard.add_log(f"📱 Sending to: {whatsapp_contact}")

                success = send_pdf_via_whatsapp(
                    driver=self.driver,
                    wait=self.wait,
                    pdf_path=merged,
                    contact_name=whatsapp_contact,
                    dashboard=self.dashboard,
                    cancel_ev=self.cancel_event,
                )

                if success:
                    self.dashboard.add_log("=" * 50)
                    self.dashboard.add_log("🎉 ALL DONE!")
                    self.dashboard.add_log("=" * 50)
                    self.dashboard.update_stats({"whatsapp_status": "sent"})
                else:
                    self.dashboard.add_log("⚠️ WhatsApp send incomplete")
                    self.dashboard.update_stats({"whatsapp_status": "failed"})

            # Give some time for WhatsApp upload to complete
            if not self.cancel_event.is_set():
                self.dashboard.add_log("⏳ Waiting 15 seconds for uploads...")
                for _ in range(15):
                    if self.cancel_event.is_set():
                        break
                    time.sleep(1)
                self.dashboard.add_log("✅ Upload wait complete")

        # Cleanup and finish
        self._finish()

    def _finish(self):
        self.safe_quit_driver()
        self.is_running = False
        self.dashboard.update_stats(
            {"status": "completed", "current_step": "-", "current_barcode": "-"}
        )
        self.dashboard.add_log("✅ Process completed")

    def _cancelled(self):
        self.safe_quit_driver()
        self.is_running = False
        self.dashboard.update_stats({"status": "idle", "current_step": "-", "current_barcode": "-"})
        self.dashboard.add_log("⏹ Cancelled by user")

    def run(self):
        # Start the dashboard server
        url = self.dashboard.start_background()
        time.sleep(1.2)

        # Open UI
        try:
            webbrowser.open(url)
        except Exception:
            pass

        print(f"\n{'=' * 60}")
        print(f"  Thyrocare Report Automation - Web Interface")
        print(f"{'=' * 60}")
        print(f"\n  Dashboard: {url}")
        print(f"  Press Ctrl+C to exit")
        print(f"\n{'=' * 60}\n")

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n\nShutting down...")
            if self.is_running:
                self.cancel_event.set()
                self.safe_quit_driver()


# ===================== Main Entry Point =====================
if __name__ == "__main__":
    app = ReportAutomation()
    app.run()