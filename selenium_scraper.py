from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import traceback
import pandas as pd
from collections import defaultdict
import time
from pathlib import Path

# Path to your ChromeDriver (if not in PATH, specify the full path)
# chromedriver_path = '/path/to/chromedriver'
# service = Service(chromedriver_path)
# driver = webdriver.Chrome(service=service)
driver = webdriver.Chrome()  # Assumes chromedriver is in PATH

# Open the target website
url = "https://portal.ecensus.gov.om"
driver.get(url)

wait = WebDriverWait(driver, 20)

# Click on 'English' link
english_xpath = '//*[@id="arabic"]/div[1]/div[1]/div[2]/button'
wait.until(EC.element_to_be_clickable((By.XPATH, english_xpath))).click()

# Click on 'Enterprises' link
enterprises_xpath = '//*[@id="english"]/div[1]/div[2]/li[3]/a'
wait.until(EC.element_to_be_clickable((By.XPATH, enterprises_xpath))).click()

time.sleep(80)  # Wait for the page to load after clicking 'Enterprises'
# Recursively collect all iframe XPaths (including nested iframes)
def collect_iframe_paths(driver, parent_chain=None):
    if parent_chain is None:
        parent_chain = []
    paths = []
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for idx, iframe in enumerate(iframes):
        # Build the frame chain for this iframe
        chain = parent_chain + [idx]
        paths.append(chain)
        try:
            driver.switch_to.frame(iframe)
            # Recursively collect nested iframes
            paths += collect_iframe_paths(driver, chain)
        except Exception as e:
            print(f"Error switching to iframe at chain {chain}: {e}")
        finally:
            driver.switch_to.default_content()
    return paths

# Collect all iframe chains
iframe_chains = collect_iframe_paths(driver)
print(f"Total iframe contexts found (including nested): {len(iframe_chains)}")
for chain in iframe_chains:
    print(f"Frame chain: {chain}")

# Try to click Establishments in each iframe context
establishment_xpath = "//span[text()='Establishments']"
for chain in iframe_chains:
    try:
        driver.switch_to.default_content()
        for idx in chain:
            driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[idx])
        print(f"Switched to frame chain: {chain}")
        elem = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, establishment_xpath)))
        print(f"Found span for Establishments in frame chain: {chain}")

        # Try ActionChains move and click
        try:
            actions = ActionChains(driver)
            actions.move_to_element(elem).click().perform()
            print(f"Clicked on Establishment using ActionChains in frame chain: {chain}")
            # === Begin: Date dropdown logic ===
            # 1. Find and click the dropdown icon
            dropdown_xpath = '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[1]/transform/div/div[2]/div/div/visual-modern/div/div/div[2]/div/i'
            try:
                dropdown_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
                actions.move_to_element(dropdown_icon).click().perform()
                print("Clicked dropdown icon.")
            # After all date options for this Establishments, break outer iframe_chains loop
            # (Place this break after the date dropdown logic block)
                time.sleep(2)
            except Exception as e:
                print(f"Could not click dropdown icon: {e}")
                break
            # 2. Find all dropdown options (look for visible options after opening dropdown)
            # This may need adjustment based on actual menu structure
            # Use Power BI slicer popup for date options
            date_option_xpath = "//div[starts-with(@id, 'slicer-dropdown-popup-')]//span"
            # Find number of options once
            options = driver.find_elements(By.XPATH, date_option_xpath)
            num_options = len(options)
            for idx in range(num_options):
                try:
                    # Re-open dropdown each time
                    dropdown_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
                    actions.move_to_element(dropdown_icon).click().perform()
                    time.sleep(1)
                    # Get fresh date options
                    options = driver.find_elements(By.XPATH, date_option_xpath)
                    if idx >= len(options):
                        continue
                    option = options[idx]
                    if option.is_displayed():
                        actions.move_to_element(option).click().perform()
                        time.sleep(10)  # Wait for data to update after date selection
                        # --- SCRAPE: Extract table-like data from two visual-modern containers ---
                        # --- Scrape Table 1 ---
                        table1_xpath = '/html/body/div[1]/ui-view/div/div/div/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[24]/transform/div/div[2]/div/div/visual-modern/div/div/div[2]/div[1]'
                        container1 = driver.find_element(By.XPATH, table1_xpath)
                        # Find all child elements within the first table container (Table 1)
                        children1 = container1.find_elements(By.XPATH, './*')
                        # Print debug information for Table 1 structure
                        print("\n[DEBUG] Table 1: Checking children for headers and rows...")
                        for idx, child in enumerate(children1):
                            print(f"[DEBUG] Child {idx}: {child.text}")

                        # --- Scrape Table 1 ---
                        # Prepare to collect all parsed rows for Table 1
                        table1_rows = []
                        # Only proceed if there is at least one data block below the header
                        if len(children1) > 1:
                            table1_data_raw = children1[1].text.split('Select Row')
                            for block in table1_data_raw:
                                lines = [line.strip().replace(',', '') for line in block.split('\n') if line.strip()]
                                if len(lines) == 4:
                                    table1_rows.append(lines)
                            if table1_rows:
                                print(f"[DEBUG] Parsed {len(table1_rows)} data rows for Table 1.")
                                for i, row in enumerate(table1_rows[:5]):
                                    print(f"[DEBUG] Table 1 Data Row {i+1}: {row}")
                                # Add Table 1 to excel_data
                                if 'excel_data' not in globals():
                                    excel_data = {}
                                date_str = option.text.strip().replace('/', '-').replace(' ', '_')
                                if not date_str:
                                    date_str = f"date_{int(time.time())}"
                                sheet_name1 = f"{date_str}_table1"
                                excel_data[sheet_name1] = {'rows': table1_rows, 'headers': ['Governorate', 'number', 'Omani', 'expats']}
                                print(f"[DEBUG] Added Table 1 to excel_data with sheet name {sheet_name1} and {len(table1_rows)} rows")

                        # --- Scrape Table 2 ---
                        # Table 2 is handled in a separate visual container, with a different column structure
                        # Prepare to collect all parsed rows for Table 2
                        table2_rows = []
                        try:
                            table2_xpath = '/html/body/div[1]/ui-view/div/div/div/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[26]/transform/div/div[2]/div/div/visual-modern/div/div/div[2]/div[1]'
                            container2 = driver.find_element(By.XPATH, table2_xpath)
                            children2 = container2.find_elements(By.XPATH, './*')
                            # Print debug information for Table 2 structure
                            print("\n[DEBUG] Table 2: Checking children for headers and rows...")
                            for idx, child in enumerate(children2):
                                print(f"[DEBUG] Child {idx}: {child.text}")
                            # Only proceed if there is at least one data block below the header for Table 2
                            if len(children2) > 1:
                                table2_data_raw = children2[1].text.split('Select Row')
                                for block in table2_data_raw:
                                    lines = [line.strip() for line in block.split('\n') if line.strip()]
                                    if len(lines) == 5:
                                        table2_rows.append(lines)
                                if table2_rows:
                                    print(f"[DEBUG] Parsed {len(table2_rows)} data rows for Table 2.")
                                    for i, row in enumerate(table2_rows[:5]):
                                        print(f"[DEBUG] Table 2 Data Row {i+1}: {row}")
                                    # Add Table 2 to excel_data
                                    if 'excel_data' not in globals():
                                        excel_data = {}
                                    date_str = option.text.strip().replace('/', '-').replace(' ', '_')
                                    if not date_str:
                                        date_str = f"date_{int(time.time())}"
                                    sheet_name2 = f"{date_str}_table2"
                                    excel_data[sheet_name2] = {'rows': table2_rows, 'headers': ['Governorate', 'school', 'masjid', 'hospital and health center', 'others']}
                                    print(f"[DEBUG] Added Table 2 to excel_data with sheet name {sheet_name2} and {len(table2_rows)} rows")
                        except Exception as e:
                            print(f"[ERROR] Table 2 scraping failed: {e}")
                        try:
                            container2 = driver.find_element(By.XPATH, other_container_xpath)
                            tables2 = container2.find_elements(By.XPATH, './/table')
                            print(f"[DEBUG] option.text for Table 2: '{option.text}'")
                            if not tables2:
                                rows2 = [line.split() for line in container2.text.split('\n') if line.strip()]
                                print(f"[DEBUG] First 5 raw rows (no tables, Table 2): {rows2[:5]}")
                                all_data[option.text + '_table2'].extend(rows2)
                            else:
                                rows2 = []
                                for t2 in tables2:
                                    for tr2 in t2.find_elements(By.XPATH, './/tr'):
                                        cells2 = [td2.text for td2 in tr2.find_elements(By.XPATH, './th|./td')]
                                        rows2.append(cells2)
                                print(f"[DEBUG] First 5 raw rows (tables, Table 2): {rows2[:5]}")
                                all_data[option.text + '_table2'].extend(rows2)
                        except Exception as e:
                            print(f"[DATA] Date: {option.text} | Other container ERROR: {e}")
                except Exception as e:
                    print(f"Could not click/select date option {idx}: {e}")
            # === End: Date dropdown logic ===
            if num_options == 0:
                print("[DATA] WARNING: No date options found in slicer.")
            else:
                print(f"[DEBUG] num_options: {num_options}")
                # Write all collected excel_data to a single Excel file with multiple sheets
                if 'excel_data' in globals() and excel_data:
                    print(f"[DEBUG] Sheets to be written: {list(excel_data.keys())}")
                    import pandas as pd
                    from pandas import ExcelWriter
                    excel_path = str(Path.home() / "Downloads" / "hihello_ALLDATES.xlsx")
                    try:
                        print(f"[DEBUG] Attempting to save Excel to {excel_path}")
                        with ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                            for i, (sheet_name, data) in enumerate(excel_data.items()):
                                safe_sheet = str(sheet_name).strip()[:31]
                                df = pd.DataFrame(data['rows'], columns=data['headers'])
                                df.to_excel(writer, sheet_name=safe_sheet, index=False)
                        print(f"[DATA] All tables saved to {excel_path}")
                    except Exception as e:
                        print(f"[DATA] ERROR saving Excel: {e}")
        except Exception as ac_e:
            print(f"ActionChains click failed in frame chain {chain}: {ac_e}")

        # Try to click the element at the span's center coordinates using JS mouse event
        try:
            rect = driver.execute_script("var r = arguments[0].getBoundingClientRect(); return {left: r.left, top: r.top, width: r.width, height: r.height};", elem)
            center_x = rect['left'] + rect['width'] / 2
            center_y = rect['top'] + rect['height'] / 2
            js = '''
                var el = document.elementFromPoint(arguments[0], arguments[1]);
                if (el) {
                    var evt = new MouseEvent('click', {
                        bubbles: true,
                        cancelable: true,
                        view: window
                    });
                    el.dispatchEvent(evt);
                    return el.outerHTML;
                } else {
                    return null;
                }
            '''
            clicked_html = driver.execute_script(js, center_x, center_y)
            print(f"Tried JS mouse event click at center ({center_x}, {center_y}). Element found: {clicked_html}")
        except Exception as jspt_e:
            print(f"JS mouse event click by point failed: {jspt_e}")

        # Still try the span click as before
        try:
            elem.click()
            print(f"Clicked on Establishment in frame chain: {chain}")
            # After all date options for this Establishments, break outer iframe_chains loop
            break
        except Exception as click_e:
            print(f"Click intercepted in frame chain {chain}: {click_e}")
            try:
                driver.execute_script("arguments[0].click();", elem)
                print(f"Clicked on Establishment using JS in frame chain: {chain}")
                # After all date options for this Establishments, break outer iframe_chains loop
                break
            except Exception as js_e:
                print(f"JS click also failed in frame chain {chain}: {js_e}")
    except Exception as e:
        print(f"Not found/clickable in frame chain {chain}: {e}")
    finally:
        driver.switch_to.default_content()

# Add any further actions below

# Optional: Close the browser after a delay
time.sleep(10)
driver.quit()
