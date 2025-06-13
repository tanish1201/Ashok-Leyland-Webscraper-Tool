import os
import time
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import glob

# --- CONFIGURATION ---
CONFIG = {
    'url': 'https://helpline.ashokleyland.com/elitesupport/',
    'login': {
        'user_field_selectors': [
            (By.CSS_SELECTOR, 'input[placeholder="Employee Id"]'),
            (By.CSS_SELECTOR, 'input[name="userId"]'),
            (By.CSS_SELECTOR, 'input[id="userId"]'),
            (By.CSS_SELECTOR, 'input[type="text"]'),
            (By.NAME, 'username'),
            (By.ID, 'username')
        ],
        'pass_field_selectors': [
            (By.CSS_SELECTOR, 'input[placeholder="Password"]'),
            (By.CSS_SELECTOR, 'input[name="password"]'),
            (By.CSS_SELECTOR, 'input[id="password"]'),
            (By.CSS_SELECTOR, 'input[type="password"]'),
            (By.NAME, 'password'),
            (By.ID, 'password')
        ],
        'submit_button_selectors': [
            (By.XPATH, "//button[normalize-space(text())='LOG IN']"),
            (By.XPATH, "//button[contains(text(), 'LOG IN')]"),
            (By.XPATH, "//button[contains(text(), 'Login')]"),
            (By.XPATH, "//input[@type='submit']"),
            (By.CSS_SELECTOR, 'button[type="submit"]'),
            (By.CSS_SELECTOR, '.login-btn'),
            (By.CSS_SELECTOR, '.btn-login')
        ]
    },
    'dashboard': {
        # Updated selectors based on the screenshot
        'date_from_selectors': [
            (By.CSS_SELECTOR, 'input[placeholder*="From"]'),
            (By.ID, 'DateFrom'),
            (By.NAME, 'dateFrom'),
            (By.CSS_SELECTOR, 'input[type="date"]'),
            (By.XPATH, "//input[contains(@placeholder, 'From') or contains(@name, 'from')]")
        ],
        'date_to_selectors': [
            (By.CSS_SELECTOR, 'input[placeholder*="To"]'),
            (By.ID, 'DateTo'),
            (By.NAME, 'dateTo'),
            (By.XPATH, "//input[contains(@placeholder, 'To') or contains(@name, 'to')]")
        ],
        'ticket_status_selectors': [
            (By.CSS_SELECTOR, 'select option[value="All"]'),
            (By.XPATH, "//select//option[text()='All']"),
            (By.CSS_SELECTOR, 'select[class*="status"] option[value="All"]')
        ],
        'zone_select_selectors': [
            (By.CSS_SELECTOR, 'select option[value="North 1"]'),
            (By.XPATH, "//select//option[text()='North 1']"),
            (By.CSS_SELECTOR, 'select[id*="zone"]')
        ],
        'region_select_selectors': [
            (By.CSS_SELECTOR, 'select option[contains(text(), "Gurgaon")]'),
            (By.CSS_SELECTOR, 'select option[contains(text(), "Delhi")]'),
            (By.XPATH, "//select//option[contains(text(), 'Gurgaon') or contains(text(), 'Delhi')]")
        ],
        'area_select_selectors': [
            (By.CSS_SELECTOR, 'select option[contains(text(), "Faridabad")]'),
            (By.CSS_SELECTOR, 'select option[contains(text(), "Gurgaon")]'),
            (By.CSS_SELECTOR, 'select option[contains(text(), "Ghaziabad")]'),
            (By.XPATH, "//select//option[contains(text(), 'Faridabad') or contains(text(), 'Gurgaon') or contains(text(), 'Ghaziabad')]")
        ],
        'dealer_select_selectors': [
            (By.CSS_SELECTOR, 'select option[contains(text(), "TTBL")]'),
            (By.XPATH, "//select//option[contains(text(), 'TTBL')]")
        ],
        'tat_select_selectors': [
            (By.CSS_SELECTOR, 'select option[value="All"]'),
            (By.XPATH, "//select//option[text()='All']")
        ],
        'submit_selectors': [
            (By.XPATH, "//button[text()='Submit']"),
            (By.CSS_SELECTOR, 'button[type="submit"]'),
            (By.XPATH, "//button[contains(text(), 'Submit')]"),
            (By.CSS_SELECTOR, 'input[type="submit"]'),
            (By.XPATH, "//input[@value='Submit']")
        ],
        'export_selectors': [
            (By.XPATH, "//button[text()='Excel']"),
            (By.CSS_SELECTOR, 'button[onclick*="excel"]'),
            (By.XPATH, "//button[contains(text(), 'Excel')]"),
            (By.ID, 'exportExcel'),
            (By.CSS_SELECTOR, 'button[id*="export"]'),
            # DataTables Excel export button selectors:
            (By.CSS_SELECTOR, 'a.buttons-excel'),
            (By.CSS_SELECTOR, 'a.exportExcel'),
            (By.CSS_SELECTOR, 'a.dt-button.buttons-excel'),
            (By.XPATH, "//a[contains(@class, 'buttons-excel')]")
        ],
        'no_data_selectors': [
            (By.XPATH, "//*[contains(text(), 'No Data Found')]"),
            (By.XPATH, "//*[contains(text(), 'No records found')]"),
            (By.XPATH, "//*[contains(text(), 'No data available')]"),
            (By.XPATH, "//*[contains(text(), 'No results')]")
        ],
        # Support type selector (top right Dealer dropdown)
        'support_type_selectors': [
            (By.CSS_SELECTOR, 'select.form-control'),  # Generic select dropdown
            (By.XPATH, "//select[contains(@class, 'form-control')]"),
            (By.CSS_SELECTOR, 'select[onchange*="support"]'),
            (By.XPATH, "//select//option[contains(text(), 'Elite') or contains(text(), 'Standard')]/..")
        ]
    }
}

# Create download directory
download_dir = os.path.join(os.getcwd(), 'downloads')
os.makedirs(download_dir, exist_ok=True)

# Clear any existing files in download directory
for file in glob.glob(os.path.join(download_dir, "*")):
    try:
        os.remove(file)
        print(f"Removed existing file: {file}")
    except:
        pass

# User credentials and mapping
users = [
    {'id': '99576WM1', 'pass': 'all@123', 'region': 'Gurgaon', 'area': 'Faridabad', 'dealer': 'TTBL Faridabad 2'},
    {'id': '99544WM1', 'pass': 'all@123', 'region': 'Gurgaon', 'area': 'Faridabad', 'dealer': 'TTBL Faridabad 1'},
    {'id': '311044WM1', 'pass': 'all@123', 'region': 'Gurgaon', 'area': 'Faridabad', 'dealer': 'TTBL Okhla'},
    {'id': '312497WM1', 'pass': 'all@123', 'region': 'Gurgaon', 'area': 'Gurgaon', 'dealer': 'TTBL Bamnoli'},
    {'id': '99577WM1', 'pass': 'all@123', 'region': 'Delhi', 'area': 'Ghaziabad', 'dealer': 'TTBL Greater Noida'},
    {'id': '312498WM1', 'pass': 'all@123', 'region': 'Gurgaon', 'area': 'Gurgaon', 'dealer': 'TTBL Gurgaon Sec18'},
    {'id': '312499WM1', 'pass': 'all@123', 'region': 'Gurgaon', 'area': 'Gurgaon', 'dealer': 'TTBL Bilaspur'}
]
# Support modes - These will be selected from the top-right dropdown
modes = [
    {'name': 'Standard Support', 'suffix': 'S', 'dropdown_text': 'Standard'},
    {'name': 'Elite Support', 'suffix': 'E', 'dropdown_text': 'Elite'}
]

# Get yesterday's date
today = datetime.now()
yesterday = (today - timedelta(days=1)).strftime('%Y-%m-%d')
yesterday_filename = (today - timedelta(days=1)).strftime('%d-%m-%Y')

print(f"Processing data for date: {yesterday}")

def setup_driver():
    """Setup Chrome WebDriver with robust options"""
    options = webdriver.ChromeOptions()
    
    prefs = {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'profile.default_content_settings.popups': 0,
        'profile.default_content_setting_values.automatic_downloads': 1
    }
    options.add_experimental_option('prefs', prefs)
    
    # Stability options
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-web-security')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    try:
        driver = webdriver.Chrome(options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.maximize_window()
        return driver
    except Exception as e:
        print(f"Error setting up Chrome driver: {e}")
        return None

def find_element_with_fallback(driver, selectors, timeout=10, clickable=False):
    """Try multiple selectors to find an element"""
    for selector in selectors:
        try:
            if clickable:
                element = WebDriverWait(driver, timeout).until(
                    EC.element_to_be_clickable(selector)
                )
            else:
                element = WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located(selector)
                )
            print(f"Found element using selector: {selector}")
            return element
        except TimeoutException:
            continue
    return None

def wait_for_download(expected_filename_part, timeout=60):
    """Wait for a file to be downloaded completely"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        for filename in os.listdir(download_dir):
            if expected_filename_part.lower() in filename.lower() and not filename.endswith('.crdownload'):
                return os.path.join(download_dir, filename)
        time.sleep(1)
    return None

def clear_and_send_keys(element, text):
    """Clear field and send keys with retry mechanism"""
    try:
        element.clear()
        time.sleep(0.5)
        element.send_keys(text)
        return True
    except Exception as e:
        print(f"Error in clear_and_send_keys: {e}")
        return False

def debug_page_source(driver, filename="debug_page.html"):
    """Save page source for debugging"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print(f"Page source saved to {filename}")
    except Exception as e:
        print(f"Could not save page source: {e}")

def check_login_success(driver, wait):
    """Check if login was successful using multiple indicators"""
    try:
        # Check URL change
        current_url = driver.current_url
        if 'consolidated-report' in current_url.lower() or 'dashboard' in current_url.lower():
            print("Login successful - redirected to dashboard/report page")
            return True
        
        # Check for the "Consolidated Report" heading
        try:
            heading = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Consolidated Report')]"))
            )
            print("Login successful - found Consolidated Report heading")
            return True
        except TimeoutException:
            pass
        
        # Check for form elements that appear after login
        dashboard_indicators = [
            (By.CSS_SELECTOR, 'input[placeholder*="From"]'),  # Date From field
            (By.CSS_SELECTOR, 'input[placeholder*="To"]'),    # Date To field
            (By.XPATH, "//button[text()='Submit']"),          # Submit button
            (By.XPATH, "//button[text()='Excel']"),           # Excel button
        ]
        
        for selector in dashboard_indicators:
            try:
                element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(selector)
                )
                print(f"Login successful - found dashboard element: {selector}")
                return True
            except TimeoutException:
                continue
        
        print("Login status unclear - no clear indicators found")
        return False
        
    except Exception as e:
        print(f"Error checking login status: {e}")
        return False

def login_user(driver, wait, user):
    """Login with improved speed and success detection"""
    try:
        print(f"Navigating to: {CONFIG['url']}")
        driver.get(CONFIG['url'])

        # Use minimal timeout since fields should be present immediately
        user_field = find_element_with_fallback(driver, CONFIG['login']['user_field_selectors'], timeout=2)
        pass_field = find_element_with_fallback(driver, CONFIG['login']['pass_field_selectors'], timeout=2)
        submit_btn = find_element_with_fallback(driver, CONFIG['login']['submit_button_selectors'], timeout=2)

        if not user_field or not pass_field or not submit_btn:
            print("Login fields/buttons not found quickly.")
            return False

        print("Entering credentials and clicking login immediately...")
        clear_and_send_keys(user_field, user['id'])
        clear_and_send_keys(pass_field, user['pass'])
        try:
            submit_btn.click()
        except Exception as e:
            print(f"Normal click failed, trying JavaScript click: {e}")
            driver.execute_script("arguments[0].click();", submit_btn)

        # Wait for dashboard element (date field or heading) after login
        dashboard_loaded = False
        dashboard_selectors = [
            (By.CSS_SELECTOR, 'input[placeholder*="From"]'),
            (By.XPATH, "//*[contains(text(), 'Consolidated Report')]")
        ]
        for selector in dashboard_selectors:
            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located(selector)
                )
                dashboard_loaded = True
                break
            except TimeoutException:
                continue
        if dashboard_loaded:
            print("Login successful - dashboard loaded")
            return True
        print("Login status unclear - dashboard not detected")
        return False
    except Exception as e:
        print(f"Login error for {user['id']}: {e}")
        return False

def select_support_type(driver, wait, mode):
    """Select support type from the top-right dropdown"""
    try:
        print(f"Selecting support type: {mode['name']}")
        
        # Look for the support type dropdown (likely in the top right area)
        support_selectors = [
            (By.XPATH, f"//select//option[contains(text(), '{mode['dropdown_text']}')]"),
            (By.CSS_SELECTOR, f'select option[value*="{mode["dropdown_text"].lower()}"]'),
            (By.XPATH, f"//option[contains(text(), '{mode['dropdown_text']}')]"),
        ]
        
        # First, try to find the select element itself
        select_element = find_element_with_fallback(driver, CONFIG['dashboard']['support_type_selectors'])
        if select_element:
            try:
                Select(select_element).select_by_visible_text(mode['dropdown_text'])
                print(f"Selected {mode['dropdown_text']} from dropdown")
                time.sleep(2)
                return True
            except Exception as e:
                print(f"Could not select from dropdown: {e}")
        
        # Alternative: try to find and click the option directly
        for selector in support_selectors:
            try:
                option = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(selector)
                )
                driver.execute_script("arguments[0].click();", option)
                print(f"Clicked {mode['dropdown_text']} option")
                time.sleep(2)
                return True
            except TimeoutException:
                continue
        
        print(f"Could not find or select {mode['dropdown_text']} support type")
        return False
        
    except Exception as e:
        print(f"Error selecting support type: {e}")
        return False

def set_form_filters(driver, wait, user, yesterday_date):
    """Set all form filters including dates, dropdowns, etc. Select ALL ticket status options."""
    try:
        print("Setting form filters...")
        # Set Date From
        print("Setting Date From...")
        date_from = find_element_with_fallback(driver, CONFIG['dashboard']['date_from_selectors'])
        if date_from:
            clear_and_send_keys(date_from, yesterday_date)
            print(f"Set Date From to: {yesterday_date}")
        else:
            print("Could not find 'Date From' field")
        # Set Date To
        print("Setting Date To...")
        date_to = find_element_with_fallback(driver, CONFIG['dashboard']['date_to_selectors'])
        if date_to:
            clear_and_send_keys(date_to, yesterday_date)
            print(f"Set Date To to: {yesterday_date}")
        else:
            print("Could not find 'Date To' field")
        # Set filters based on user data with proper sequencing
        try:
            # Set Zone first and wait for Region dropdown
            print("Setting Zone to North 1...")
            zone_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "zone"))
            )
            Select(zone_select).select_by_visible_text("North 1")
            # Wait for Region dropdown to be populated
            region_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "state"))
            )
            WebDriverWait(driver, 10).until(
                lambda d: len(Select(region_select).options) > 1
            )
            Select(region_select).select_by_visible_text(user['region'])
            # Wait for Area dropdown to be populated
            area_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "city"))
            )
            WebDriverWait(driver, 10).until(
                lambda d: len(Select(area_select).options) > 1
            )
            Select(area_select).select_by_visible_text(user['area'])
            # Wait for Dealer dropdown to be populated
            dealer_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "dealer"))
            )
            found_dealer = False
            for i in range(20):
                select_obj = Select(dealer_select)
                dealer_options = [opt.text for opt in select_obj.options]
                print(f"[Wait {i+1}s] Dealer options: {dealer_options}")
                if user['dealer'] in dealer_options:
                    found_dealer = True
                    break
                time.sleep(0.5)  # Reduced to minimum for polling
            if not found_dealer:
                print(f"Dealer '{user['dealer']}' not found in dropdown after 10 seconds.")
                return False
            try:
                select_obj.select_by_visible_text(user['dealer'])
                print(f"Successfully set Dealer to {user['dealer']} (visible text match)")
            except Exception as e:
                print(f"select_by_visible_text failed: {e}")
                return False
        except Exception as e:
            print(f"Error in dealer selection: {e}")
            return False

        except Exception as e:
            print(f"Error in filter selection sequence: {e}")
            return False

        # Select ALL Ticket Status options (multi-select)
        print("Selecting ALL Ticket Status options...")
        try:
            ticket_status = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ticketStatus"))
            )
            select = Select(ticket_status)
            for option in select.options:
                select.select_by_visible_text(option.text)
            print("Selected all Ticket Status options")
        except Exception as e:
            print(f"Error selecting Ticket Status options: {e}")

        # Set TAT to All
        print("Setting TAT to All...")
        try:
            tat_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "tat"))
            )
            Select(tat_select).select_by_visible_text("All")
            print("Set TAT to All")
        except Exception as e:
            print(f"Error setting TAT: {e}")

        time.sleep(2)
        return True
    except Exception as e:
        print(f"Error setting form filters: {e}")
        return False

def get_current_support_mode(driver):
    """Detects the current support mode by reading the top-middle heading ('Elite Support' or 'Standard Support')."""
    try:
        # Try to find the heading in h1, h2, h3, h4, or .card-title
        heading_selectors = [
            (By.XPATH, "//h1[contains(text(), 'Support') or contains(text(), 'support') or contains(text(), 'SUPPORT')]") ,
            (By.XPATH, "//h2[contains(text(), 'Support') or contains(text(), 'support') or contains(text(), 'SUPPORT')]") ,
            (By.XPATH, "//h3[contains(text(), 'Support') or contains(text(), 'support') or contains(text(), 'SUPPORT')]") ,
            (By.XPATH, "//h4[contains(text(), 'Support') or contains(text(), 'support') or contains(text(), 'SUPPORT')]") ,
            (By.CSS_SELECTOR, ".card-title"),
        ]
        for selector in heading_selectors:
            try:
                elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located(selector))
                text = elem.text.strip().lower()
                if 'elite' in text:
                    return 'Elite Support'
                elif 'standard' in text:
                    return 'Standard Support'
            except Exception:
                continue
        # Fallback: check page source
        page = driver.page_source.lower()
        if 'elite support' in page:
            return 'Elite Support'
        elif 'standard support' in page:
            return 'Standard Support'
        return 'Unknown'
    except Exception as e:
        print(f"Error detecting support mode: {e}")
        return 'Unknown'

def switch_support_mode(driver, wait, target_mode):
    """Switch to the target support mode by clicking Dealer and selecting the other support."""
    try:
        current_mode = get_current_support_mode(driver)
        print(f"Current support mode: {current_mode}")
        print(f"Target support mode: {target_mode['name']}")
        if current_mode == target_mode['name']:
            print(f"Already on {target_mode['name']}")
            return True
        # Click Dealer button (top right)
        try:
            # Most robust: by id
            dealer_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@id='profileDropdown']"))
            )
            driver.execute_script("arguments[0].click();", dealer_btn)
            time.sleep(1)
        except Exception:
            # Fallback: by class
            try:
                dealer_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a#profileDropdown"))
                )
                driver.execute_script("arguments[0].click();", dealer_btn)
                time.sleep(1)
            except Exception as e:
                print(f"Could not find Dealer button: {e}")
                return False
        # Now, in the dropdown, click the correct support mode
        try:
            support_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//a[contains(@class, 'dropdown-item') and contains(., '{target_mode['dropdown_text']}')]"))
            )
            driver.execute_script("arguments[0].click();", support_link)
            print(f"Clicked to switch to {target_mode['name']}")
            time.sleep(5)
            for _ in range(10):
                if get_current_support_mode(driver) == target_mode['name']:
                    print(f"Switched to {target_mode['name']}")
                    return True
                time.sleep(1)
            print(f"Failed to switch to {target_mode['name']} after clicking")
            return False
        except Exception as e:
            print(f"Could not find support mode link in Dealer dropdown: {e}")
            return False
    except Exception as e:
        print(f"Error in switch_support_mode: {e}")
        return False

def process_user_mode(driver, wait, user, mode):
    """Process a single user for a specific support mode, selecting all ticket status options at once."""
    try:
        print(f"Processing {user['dealer']} - {mode['name']}")
        if not switch_support_mode(driver, wait, mode):
            print(f"Failed to switch to {mode['name']}, trying to continue anyway...")
        time.sleep(3)
        if not set_form_filters(driver, wait, user, yesterday):
            print("Failed to set some filters, continuing anyway...")
            return None
        print("Submitting form...")
        submit_btn = find_element_with_fallback(driver, CONFIG['dashboard']['submit_selectors'], timeout=10, clickable=True)
        if submit_btn:
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
                submit_btn.click()
                print("Form submitted successfully")
            except Exception as e:
                print(f"Error clicking submit button: {e}")
                return None
        else:
            print("Could not find submit button")
            return None

        # Wait for table to load before looking for Excel button
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table tbody tr'))
            )
            print("Table loaded.")
        except TimeoutException:
            print("Table did not load in time.")

        # Now look for the Excel button with a short timeout
        print("Looking for Excel export button...")
        export_btn = find_element_with_fallback(driver, CONFIG['dashboard']['export_selectors'], timeout=5, clickable=True)
        if not export_btn:
            print(f"Excel export button not found quickly.")
            debug_page_source(driver, f"no_excel_button_{user['id']}_{mode['suffix']}.html")
            return None
        try:
            print("Clicking Excel export button...")
            driver.execute_script("arguments[0].scrollIntoView(true);", export_btn)
            time.sleep(2)
            files_before = set(os.listdir(download_dir))
            driver.execute_script("arguments[0].click();", export_btn)
            print("Excel export button clicked")
            print("Waiting for download to complete...")
            start_time = time.time()
            downloaded_file = None
            while time.time() - start_time < 45:
                files_after = set(os.listdir(download_dir))
                new_files = files_after - files_before
                for fname in new_files:
                    if fname.endswith('.xlsx') and not fname.endswith('.crdownload'):
                        downloaded_file = os.path.join(download_dir, fname)
                        break
                if downloaded_file:
                    break
                time.sleep(1)
            if not downloaded_file:
                print(f"Download failed for {user['dealer']} - {mode['name']}")
                return None
            print(f"Download completed: {os.path.basename(downloaded_file)}")
            dealer_name_clean = user['dealer'].replace(' ', '_').replace('/', '_')
            new_filename = f"{dealer_name_clean}_{yesterday_filename}_{mode['suffix']}_ALL_TICKET_STATUS.xlsx"
            new_filepath = os.path.join(download_dir, new_filename)
            try:
                os.rename(downloaded_file, new_filepath)
                print(f"File renamed to: {new_filename}")
                return [new_filepath]
            except Exception as e:
                print(f"Error renaming file: {e}")
                return [downloaded_file]
        except Exception as e:
            print(f"Error during download process: {e}")
            return None
    except Exception as e:
        print(f"Error processing {user['dealer']} - {mode['name']}: {e}")
        return None

def main():
    """Main execution function"""
    downloaded_files = []
    for user in users:
        print(f"\n{'='*50}")
        print(f"Processing user: {user['id']} - {user['dealer']}")
        print(f"{'='*50}")
        driver = setup_driver()
        if not driver:
            print("Failed to setup Chrome driver, skipping user.")
            continue
        wait = WebDriverWait(driver, 20)
        try:
            if not login_user(driver, wait, user):
                print(f"Login failed for {user['id']}, skipping to next user.")
                driver.quit()
                continue
            # Always process Elite Support first, then Standard Support
            for mode in sorted(modes, key=lambda m: 0 if m['name'] == 'Elite Support' else 1):
                files = process_user_mode(driver, wait, user, mode)
                if files:
                    downloaded_files.extend(files)
                time.sleep(3)
            print(f"Completed processing for {user['id']}")
        except Exception as e:
            print(f"Unexpected error for user {user['id']}: {e}")
        finally:
            driver.quit()
    # Combine files
    if downloaded_files:
        print(f"\n{'='*50}")
        print("Combining all downloaded files into separate sheets...")
        print(f"{'='*50}")
        import openpyxl
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        output_filename = f"Combined_Report_{yesterday_filename}.xlsx"
        output_path = os.path.join(os.getcwd(), output_filename)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for file_path in downloaded_files:
                try:
                    print(f"Reading: {os.path.basename(file_path)}")
                    df = pd.read_excel(file_path)
                    sheet_name = os.path.splitext(os.path.basename(file_path))[0][:31]  # Excel sheet name max 31 chars
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    print(f"Error reading {file_path}: {e}")
        print(f"\nCombined file saved as: {output_filename}")
        print(f"Files combined as separate sheets: {len(downloaded_files)}")
    else:
        print("No files were downloaded successfully.")

if __name__ == "__main__":
    main()
