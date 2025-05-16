from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import json
import time
import os

def get_page_range():
    while True:
        try:
            first_page = int(input("Enter the first page number to scrape: ").strip())
            if first_page < 1:
                print("First page must be at least 1.")
                continue
                
            last_page = int(input("Enter the last page number to scrape: ").strip())
            if last_page < first_page:
                print("Last page must be greater than or equal to first page.")
                continue
                
            return first_page, last_page
        except ValueError:
            print("Please enter valid integer values for page numbers.")

def extract_page_data(driver, page_number, folder="assessment_reports"):
    print(f"Extracting data from page {page_number}...")
    soup = BeautifulSoup(driver.page_source, "html.parser")
    table = soup.find("table", id="MainContent_gvReferences")
    
    if not table:
        print("Table not found.")
        return False
        
    page_data = []
    
    # Get all rows
    rows = table.find_all("tr")
    
    # Apply more comprehensive filtering for pagination rows
    filtered_rows = []
    for row in rows:
        # Skip rows with class 'pgr'
        if 'pgr' in row.get('class', []):
            continue
        # Skip rows containing pagination links
        if row.find('a', href=lambda x: x and 'Page$' in x):
            continue
        # Skip rows containing page number spans
        if row.find('span') and row.find('span').text.strip().isdigit():
            continue
        filtered_rows.append(row)
    
    # Skip header row if present
    if len(filtered_rows) > 1:
        filtered_rows = filtered_rows[1:]
    
    for idx, row in enumerate(filtered_rows, start=1):
        cols = row.find_all("td")
        if len(cols) >= 6:
            download_link = None
            if cols[5].find("a"):
                download_link = "https://app.nwtgeoscience.ca " + cols[5].find("a")["href"]

            page_data.append({
                "S.No": idx-2,
                "Reference": cols[0].text.strip(),
                "Type": cols[1].text.strip(),
                "Title": cols[2].text.strip(),
                "Company": cols[3].text.strip(),
                "Date": cols[4].text.strip(),
                "Location": cols[5].text.strip(),
                "Download Link": download_link
            })

    filename = f"PAGE_{page_number}.json"
    filepath = os.path.join(folder, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(page_data, f, indent=2)
    print(f"Saved {len(page_data)} records to {filepath}")
    return True

def navigate_to_page(driver, target_page):
    """Navigate to a specific page number"""
    current_page_element = driver.find_element(By.CSS_SELECTOR, "span.current")
    current_page = int(current_page_element.text.strip())
    
    print(f"Currently on page {current_page}, navigating to page {target_page}")
    
    if current_page == target_page:
        return True
    
    # Try direct navigation if possible
    try:
        page_link_xpath = f"//a[contains(@href, \"javascript:__doPostBack('ctl00$MainContent$gvReferences','Page${target_page}')\")]"
        page_link = driver.find_element(By.XPATH, page_link_xpath)
        driver.execute_script("arguments[0].scrollIntoView();", page_link)
        time.sleep(1)
        page_link.click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
        time.sleep(2)
        return True
    except:
        # If direct navigation fails, try incremental navigation
        while current_page < target_page:
            try:
                next_link = driver.find_element(By.XPATH, "//a[contains(@href, 'Page$Next')]")
                driver.execute_script("arguments[0].scrollIntoView();", next_link)
                time.sleep(1)
                next_link.click()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
                time.sleep(2)
                current_page_element = driver.find_element(By.CSS_SELECTOR, "span.current")
                current_page = int(current_page_element.text.strip())
            except Exception as e:
                print(f"Navigation error: {str(e)}")
                return False
        return current_page == target_page

def scrape_pages_in_range(driver, first_page, last_page, folder="assessment_reports"):
    """Scrape only the pages within the specified range"""
    os.makedirs(folder, exist_ok=True)  # Create folder if it doesn't exist
    
    # Navigate to the first page if not already there
    if first_page > 1:
        if not navigate_to_page(driver, first_page):
            print(f"Failed to navigate to page {first_page}")
            return
    
    current_page = first_page
    
    while current_page <= last_page:
        if not extract_page_data(driver, current_page, folder):
            print("No data extracted. Stopping.")
            break
            
        if current_page == last_page:
            print(f"Reached specified last page ({last_page}).")
            break

        try:
            # Try to find next page link
            next_page_number = current_page + 1
            next_page_xpath = f"//a[contains(@href, \"javascript:__doPostBack('ctl00$MainContent$gvReferences','Page${next_page_number}')\")]"
            
            try:
                next_link = driver.find_element(By.XPATH, next_page_xpath)
            except:
                next_link = driver.find_element(By.XPATH, "//a[contains(@href, 'Page$Next') or contains(text(), '...')]")

            driver.execute_script("arguments[0].scrollIntoView();", next_link)
            time.sleep(1)
            next_link.click()
            
            # Wait for table to reload
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
            time.sleep(2)
            current_page += 1
        except Exception as e:
            print("No more pages or encountered an error:", str(e))
            break

def main():
    # Setup Chrome options
    options = Options()
    options.add_argument("--headless")  # Remove this line if you want to see browser UI
    options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(options=options)

    try:
        # Open website
        print("Opening website...")
        driver.get("https://app.nwtgeoscience.ca/Searching/ReferenceSearch.aspx ")
        time.sleep(3)

        # Select Assessment Report type
        print("Selecting Assessment Report type...")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_butReferenceType"))).click()
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_gvRefTypeSel")))
        time.sleep(2)

        # Find and select Assessment Report option
        rows = driver.find_element(By.ID, "MainContent_gvRefTypeSel").find_elements(By.TAG_NAME, "tr")
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 2 and "Assessment Report" in cells[1].text:
                checkbox = cells[0].find_element(By.TAG_NAME, "input")
                driver.execute_script("arguments[0].click();", checkbox)
                print("'Assessment Report' selected")
                break
        else:
            raise Exception("Assessment Report option not found")

        # Apply selection and search
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_btnApplyRefType"))).click()
        time.sleep(2)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_btnSearch"))).click()

        # Wait for results and start scraping
        print("Waiting for results...")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "MainContent_gvReferences")))
        print("Table loaded. Starting scraping...")
        
        # Get page range from user
        first_page, last_page = get_page_range()
        
        # Use the new function to scrape pages in the specified range
        scrape_pages_in_range(driver, first_page, last_page)
        print(f"Pages {first_page} to {last_page} scraped.")

    except Exception as e:
        print(f"Error: {str(e)}")
        with open("error_page.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("Page source saved as error_page.html for debugging.")
    finally:
        driver.quit()
        print("Driver closed.")

if __name__ == "__main__":
    main()
