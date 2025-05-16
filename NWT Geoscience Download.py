from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import re

def get_user_input(prompt):
    while True:
        response = input(f"{prompt} (yes/no): ").lower().strip()
        if response in ["yes", "y"]:
            return True
        elif response in ["no", "n"]:
            return False
        print("Please enter yes or no.")

def get_page_range():
    while True:
        try:
            first_page = int(input("Enter the first page number to download: ").strip())
            if first_page < 1:
                print("First page must be at least 1.")
                continue
            last_page = int(input("Enter the last page number to download: ").strip())
            if last_page < first_page:
                print("Last page must be greater than or equal to first page.")
                continue
            return first_page, last_page
        except ValueError:
            print("Please enter valid integer values for page numbers.")

def wait_for_downloads(download_dir, timeout=18000):
    """Wait for all Chrome downloads to complete. Return elapsed time in seconds."""
    print(f"Waiting for downloads to complete (timeout: {timeout}s)...")
    start_overall = time.time()

    # Wait for any .crdownload to appear
    start_detect = time.time()
    while time.time() - start_detect < timeout:
        if any(fn.endswith('.crdownload') for fn in os.listdir(download_dir)):
            break
        time.sleep(0.5)
    else:
        print("No downloads appear to have started. Continuing anyway...")
        return 0.0

    # Wait for all .crdownload to disappear
    while True:
        if not any(fn.endswith('.crdownload') for fn in os.listdir(download_dir)):
            elapsed = time.time() - start_overall
            print(f"All downloads completed in {elapsed:.1f} seconds.")
            time.sleep(2)
            return elapsed
        if time.time() - start_overall > timeout:
            print(f"Download timeout after {timeout} seconds.")
            return timeout
        time.sleep(1)

def navigate_to_page(driver, target_page):
    try:
        current = driver.find_element(By.CSS_SELECTOR, "span.current").text.strip()
        current_page = int(current)
    except:
        current_page = 1

    print(f"Currently on page {current_page}, navigating to page {target_page}")
    if current_page == target_page:
        return True

    # Try direct jump
    try:
        xpath = f"//a[contains(@href, \"javascript:__doPostBack('ctl00$MainContent$gvReferences','Page${target_page}')\")]"
        link = driver.find_element(By.XPATH, xpath)
        driver.execute_script("arguments[0].scrollIntoView();", link)
        time.sleep(1)
        link.click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
        time.sleep(2)
        return True
    except:
        # Fallback: click “Next” until you reach it
        while current_page < target_page:
            try:
                nxt = driver.find_element(By.XPATH, "//a[contains(@href, 'Page$Next')]")
                driver.execute_script("arguments[0].scrollIntoView();", nxt)
                time.sleep(1)
                nxt.click()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
                time.sleep(2)
                try:
                    current_page = int(driver.find_element(By.CSS_SELECTOR, "span.current").text.strip())
                except:
                    current_page += 1
            except Exception as e:
                print(f"Navigation error: {e}")
                return False
        return True

def find_download_links(driver):
    try:
        print("Scanning for download links...")
        table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))

        # Method 1: icons
        icons = driver.find_elements(By.CSS_SELECTOR, "i.fa-download")
        links = []
        for icon in icons:
            try:
                a = icon.find_element(By.XPATH, "./..")
                if a.tag_name == 'a' and a.get_attribute('href'):
                    links.append(a)
            except:
                pass
        if links:
            print(f"Found {len(links)} via icon")
            return [(i+1, link) for i, link in enumerate(links)]

        # Method 2: href pattern
        all_a = driver.find_elements(By.TAG_NAME, "a")
        links = [a for a in all_a if a.get_attribute('href') and 'DownloadRec' in a.get_attribute('href')]
        if links:
            print(f"Found {len(links)} via href pattern")
            return [(i+1, link) for i, link in enumerate(links)]

        # Method 3: scan rows
        rows = table.find_elements(By.TAG_NAME, "tr")
        links = []
        for idx, row in enumerate(rows[1:], start=1):
            try:
                a_tags = row.find_elements(By.TAG_NAME, "a")
                for a in a_tags:
                    if a.get_attribute('href') and ('DownloadRec' in a.get_attribute('href') or '__doPostBack' in a.get_attribute('href')):
                        links.append((idx, a))
                        break
            except:
                pass
        if links:
            print(f"Found {len(links)} via table scan")
            return links

        print("No download links found.")
        return []
    except Exception as e:
        print(f"Error scanning links: {e}")
        return []

def get_row_numbers_to_download():
    while True:
        txt = input("Enter row number(s) (e.g., '1,3,5-7' or 'all'): ").strip().lower()
        if txt == 'all':
            return None
        try:
            nums = set()
            for part in txt.split(','):
                part = part.strip()
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    nums.update(range(start, end+1))
                else:
                    nums.add(int(part))
            return sorted(nums)
        except:
            print("Invalid format. Try again.")

def download_individual_reports(driver, download_dir):
    print("Starting individual downloads...")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
    links = find_download_links(driver)
    if not links:
        driver.save_screenshot("no_links.png")
        with open("no_links.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        return False

    row_nums = get_row_numbers_to_download()
    link_dict = {i: link for i, link in links}
    to_download = sorted(link_dict.keys()) if row_nums is None else [n for n in row_nums if n in link_dict]

    if row_nums:
        missing = [n for n in row_nums if n not in link_dict]
        if missing:
            print(f"Rows {missing} not found; skipping them.")

    print(f"Will download rows: {to_download}")

    for idx, row in enumerate(to_download):
        # refetch links to avoid stale elements
        if idx > 0:
            links = find_download_links(driver)
            link_dict = {i: link for i, link in links}
            if row not in link_dict:
                print(f"Row {row} disappeared; skipping.")
                continue

        link = link_dict[row]
        print(f"Downloading row {row}...")
        driver.execute_script("arguments[0].scrollIntoView(true);", link)
        time.sleep(1)

        # try click methods
        success = False
        for method in (lambda e: e.click(),
                       lambda e: driver.execute_script("arguments[0].click();", e),
                       None):
            if method:
                try:
                    method(link)
                    success = True
                    break
                except:
                    pass
            else:
                # last resort: exec JS href
                href = link.get_attribute('href')
                if href and href.startswith('javascript:'):
                    driver.execute_script(href.replace('javascript:', ''))
                    success = True
                    break

        if not success:
            print(f"Failed to click row {row}")
            continue

        # wait for download dialog
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_ASPxButton1")))
            time.sleep(2)
            driver.find_element(By.ID, "MainContent_ASPxButton1").click()
        except:
            print("Could not click 'Download All'")
        elapsed = wait_for_downloads(download_dir, timeout=18000)

        # remove all
        try:
            btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "MainContent_ASPxButton3")))
            btn.click()
            time.sleep(2)
        except:
            pass
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
        time.sleep(2)

        # decide whether to prompt
        if elapsed >= 18000 and idx != len(to_download)-1:
            if not get_user_input("Download hit timeout. Continue to next?"):
                print("Stopping as per user request.")
                break
        else:
            print("Proceeding automatically to next report...")

    return True

def download_page_reports(driver, page, download_dir):
    print(f"\n--- Page {page} ---")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
    return download_individual_reports(driver, download_dir)

def download_pages_in_range(driver, first_page, last_page, download_dir):
    if first_page > 1 and not navigate_to_page(driver, first_page):
        print(f"Could not go to page {first_page}")
        return

    curr = first_page
    while curr <= last_page:
        if not download_page_reports(driver, curr, download_dir):
            print("Halting due to an error.")
            break
        if curr == last_page:
            print(f"Done through page {last_page}.")
            break
        if not get_user_input(f"Proceed to page {curr+1}?"):
            print("User halted at page transition.")
            break
        # click Next
        try:
            xpath = f"//a[contains(@href, \"Page${curr+1}\")]"
            nxt = driver.find_element(By.XPATH, xpath)
            driver.execute_script("arguments[0].scrollIntoView();", nxt)
            time.sleep(1)
            nxt.click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
            time.sleep(2)
            curr += 1
        except Exception as e:
            print(f"Could not advance: {e}")
            break

def main():
    options = Options()
    options.add_argument("--window-size=1920,1080")
    download_dir = os.path.join(os.getcwd(), "downloaded_reports")
    os.makedirs(download_dir, exist_ok=True)
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    for arg in ("--disable-extensions","--disable-gpu","--no-sandbox","--disable-dev-shm-usage"):
        options.add_argument(arg)

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(1000)

    try:
        driver.get("https://app.nwtgeoscience.ca/Searching/ReferenceSearch.aspx")
        time.sleep(3)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_butReferenceType"))).click()
        time.sleep(2)
        rows = driver.find_element(By.ID, "MainContent_gvRefTypeSel").find_elements(By.TAG_NAME, "tr")
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols)>=2 and "Assessment Report" in cols[1].text:
                driver.execute_script("arguments[0].click();", cols[0].find_element(By.TAG_NAME, "input"))
                break
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_btnApplyRefType"))).click()
        time.sleep(2)
        driver.find_element(By.ID, "MainContent_btnSearch").click()
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))

        print("\nDownload Mode:\n1) Pick pages & rows\n2) All pages\n")
        while True:
            mode = input("Choice (1/2): ").strip()
            if mode in ('1','2'): break

        if mode=='1':
            fp, lp = get_page_range()
            download_pages_in_range(driver, fp, lp, download_dir)
        else:
            auto = get_user_input("Auto-advance through all pages without prompt?")
            download_pages_in_range(driver, 1, 9999, download_dir)

        print("All done.")

    except Exception as e:
        print("Fatal error:", e)
        try:
            driver.save_screenshot("error.png")
            with open("error.html","w",encoding="utf-8") as f:
                f.write(driver.page_source)
        except:
            pass
    finally:
        driver.quit()
        print("Driver closed.")

if __name__ == "__main__":
    main()
