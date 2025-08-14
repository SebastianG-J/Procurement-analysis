"""
Meter per Roll Web Scraper - Documented Version
===============================================

This script automates the collection of product specifications from supplier websites,
specifically focusing on "meter per roll" data and base units for roll-based products.

Purpose:
- Scrape product specifications from supplier websites
- Extract meter per roll information for roll-based products
- Collect base unit information (MTR, Mtr.)
- Filter and process product lists to avoid duplicate scraping

Key Features:
- Automated web scraping using Selenium
- Periodic data saving to prevent data loss
- Intelligent filtering of already scraped products
- Robust error handling and timeout management

Author: [Your Name]
Date: [Date]
"""

import argparse
import logging
import os
import time
import random
import threading

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ===== XPATH SELECTORS =====
# These XPath selectors target specific elements on the supplier website
SEARCH_INPUT_XPATH = "/html/body/header/div/div[1]/div[1]/div/form/input"
DROPDOWN_FIRST_ITEM_XPATH = "/html/body/header/div/div[1]/div[1]/div/div/a[1]"
COOKIE_ACCEPT_XPATH = '//*[@id="coiPage-1"]/div[2]/div[1]/button[3]'
DEFAULT_TABLE_XPATH = '//*[@id="addMultipleToCartForm"]/div/table'

# ===== LOGGING CONFIGURATION =====
# Set up logging to track the scraping progress and any issues
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# ===== WEB DRIVER INITIALIZATION =====
def init_driver(headless: bool):
    """
    Initialize and configure the Chrome web driver for web scraping.
    
    Args:
        headless (bool): Whether to run the browser in headless mode (no GUI)
    
    Returns:
        webdriver.Chrome: Configured Chrome driver instance
    """
    options = webdriver.ChromeOptions()
    if headless:
        # Run browser without GUI for server environments
        options.add_argument("--headless=new")
    
    # Disable automation detection to avoid being blocked
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Set window size for consistent element positioning
    options.add_argument("--window-size=1280,800")
    
    # Create and configure the driver
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(1)  # Wait up to 1 second for elements to appear
    return driver

# ===== EXCEL DATA LOADING =====
def load_varenr(path: str):
    """
    Load product numbers (Varenr.) from an Excel file.
    
    Args:
        path (str): Path to the Excel file containing product numbers
    
    Returns:
        set: Set of unique product numbers
    
    Raises:
        FileNotFoundError: If the Excel file doesn't exist
        KeyError: If the 'Varenr.' column is not found
    """
    # Check if the file exists
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Excel file not found: {path}")
    
    # Load the Excel file
    df = pd.read_excel(path, engine="openpyxl")
    
    # Verify the required column exists
    if "Varenr." not in df.columns:
        raise KeyError(f"'Varenr.' column not found. Found: {list(df.columns)}")
    
    # Clean and process the product numbers
    values = df["Varenr."].dropna().astype(str).str.strip()
    values = values[values != ""].drop_duplicates()
    return set(values)

# ===== COOKIE HANDLING =====
def accept_cookies(driver):
    """
    Accept cookies on the website to avoid popup interference.
    
    Args:
        driver: Chrome web driver instance
    """
    try:
        # Wait for cookie accept button to be clickable
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, COOKIE_ACCEPT_XPATH))
        )
        btn.click()
        
        # Wait for cookie popup to disappear
        WebDriverWait(driver, 5).until(
            EC.invisibility_of_element_located((By.XPATH, COOKIE_ACCEPT_XPATH))
        )
    except TimeoutException:
        # Cookie popup might not appear, continue silently
        pass
    except Exception:
        # Any other error, continue silently
        pass

# ===== PRODUCT SEARCH =====
def search_and_open_product(driver, varenr):
    """
    Search for a product on the website and open its details page.
    
    Args:
        driver: Chrome web driver instance
        varenr (str): Product number to search for
    
    Returns:
        bool: True if product was found and opened, False otherwise
    """
    try:
        # Wait for search input to be available and clickable
        search_input = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, SEARCH_INPUT_XPATH))
        )
        
        # Clear any existing text and enter the product number
        search_input.clear()
        search_input.send_keys(varenr)
        
        # Wait for and click the first search result
        first_item = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, DROPDOWN_FIRST_ITEM_XPATH))
        )
        first_item.click()
        
        # Brief pause to let the page load
        time.sleep(0.5)
        return True
        
    except TimeoutException:
        # Product not found or search failed
        return False

# ===== TABLE PARSING =====
def normalize_header(text: str) -> str:
    """
    Normalize table header text for consistent matching.
    
    Args:
        text (str): Raw header text
    
    Returns:
        str: Normalized header text (lowercase, single spaces)
    """
    return " ".join(text.lower().split())

def build_col_map(driver, table_xpath):
    """
    Build a mapping of column headers to their indices in the table.
    
    Args:
        driver: Chrome web driver instance
        table_xpath (str): XPath to the table element
    
    Returns:
        dict: Mapping of normalized header names to column indices
    """
    # Try to find headers in the thead section first
    headers = driver.find_elements(By.XPATH, f"{table_xpath}//thead//th")
    
    # If no thead headers, look for headers in the first row of tbody
    if not headers:
        headers = driver.find_elements(
            By.XPATH,
            f"{table_xpath}//tbody/tr[1]/th | {table_xpath}//tbody/tr[1]/td"
        )
    
    # Build the column mapping
    col_map = {}
    for idx, h in enumerate(headers):
        txt = h.text.strip()
        if txt:
            col_map[normalize_header(txt)] = idx
    return col_map

def extract_metrics(driver, varenr, table_xpath=DEFAULT_TABLE_XPATH):
    """
    Extract meter per roll and base unit information from the product table.
    
    Args:
        driver: Chrome web driver instance
        varenr (str): Product number to extract data for
        table_xpath (str): XPath to the table containing product data
    
    Returns:
        dict: Dictionary containing 'meter_pr_rulle' and 'basisenhed' values
    """
    # Initialize result with empty values
    result = {"meter_pr_rulle": "", "basisenhed": ""}
    
    # Build column mapping for the table
    col_map = build_col_map(driver, table_xpath)
    
    # Define which columns we want to extract
    desired = {
        "meter pr. rulle": "meter_pr_rulle",
        "basisenhed": "basisenhed",
    }

    # Try to find the row containing the product number
    try:
        # Exact match first
        row = driver.find_element(
            By.XPATH,
            f"{table_xpath}//tbody/tr[.//td[normalize-space()='{varenr}']]"
        )
    except:
        try:
            # Partial match if exact match fails
            row = driver.find_element(
                By.XPATH,
                f"{table_xpath}//tbody/tr[.//td[contains(normalize-space(.), '{varenr}')]]"
            )
        except:
            # Product not found in table
            return result

    # Extract data from the found row
    cells = row.find_elements(By.XPATH, ".//td")
    
    # Extract values for each desired column
    for header, key in desired.items():
        idx = col_map.get(header)
        if idx is not None and idx < len(cells):
            result[key] = cells[idx].text.strip()
        else:
            # Fallback to hardcoded indices if column mapping fails
            if key == "meter_pr_rulle" and len(cells) > 4:
                result[key] = cells[4].text.strip()
            elif key == "basisenhed" and len(cells) > 6:
                result[key] = cells[6].text.strip()

    # Clean up and validate the extracted data
    # Check if meter_pr_rulle is a valid number
    if not result["meter_pr_rulle"].replace(",", ".").replace(".", "", 1).isdigit():
        result["meter_pr_rulle"] = ""
    
    # Check if basisenhed is a valid unit
    if result["basisenhed"] not in ["MTR", "Mtr."]:
        result["basisenhed"] = ""
    
    return result

# ===== DATA SAVING =====
def write_excel(results, path):
    """
    Save the scraped results to an Excel file.
    
    Args:
        results (list): List of dictionaries containing scraped data
        path (str): Path where the Excel file should be saved
    """
    if results:
        pd.DataFrame(results).to_excel(path, index=False)
        # Note: Logging statement removed to avoid exposing confidential data

# ===== PERIODIC SAVING =====
def periodic_save(results, path, stop_event):
    """
    Periodically save results to prevent data loss during long scraping sessions.
    
    Args:
        results (list): List of scraped results (shared between threads)
        path (str): Path where to save the data
        stop_event (threading.Event): Event to signal when to stop saving
    """
    while not stop_event.is_set():
        time.sleep(60)  # Save every 60 seconds
        write_excel(results, path)
        # Note: Logging statement removed to avoid exposing confidential data

# ===== MAIN EXECUTION =====
def main():
    """
    Main function that orchestrates the entire scraping process.
    """
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description="Scrape product specifications from supplier website.")
    parser.add_argument("--search-url", default="https://www.alfotech.dk/", help="Search page URL")
    parser.add_argument("--full-list", default="vareposter_antal_salg_køb_modregning_opdateret_med_købt.xlsx", help="Full product list")
    parser.add_argument("--exclude-xlsx", default="results_merged.xlsx", help="Already scraped list")
    parser.add_argument("--output-xlsx", default="results_new.xlsx", help="New output file")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    args = parser.parse_args()

    # Load product lists
    full_list = load_varenr(args.full_list)
    exclude_list = load_varenr(args.exclude_xlsx)
    
    # Calculate which products need to be scraped
    to_scrape = sorted(full_list - exclude_list)

    # Filter out products starting with "25" (specific business rule)
    to_scrape = [v for v in to_scrape if not str(v).startswith("25")]

    # Note: Logging statement removed to avoid exposing confidential data

    # Initialize web driver and navigate to search page
    driver = init_driver(headless=args.headless)
    driver.get(args.search_url)
    accept_cookies(driver)

    # Initialize results list and periodic saving
    results = []
    stop_event = threading.Event()
    save_thread = threading.Thread(target=periodic_save, args=(results, args.output_xlsx, stop_event))
    save_thread.daemon = True
    save_thread.start()

    # Scrape each product
    for varenr in to_scrape:
        # Note: Logging statement removed to avoid exposing confidential data
        
        # Initialize entry for this product
        entry = {"Varenr.": varenr, "meter_pr_rulle": "", "basisenhed": ""}
        
        # Try to scrape the product data
        if search_and_open_product(driver, varenr):
            metrics = extract_metrics(driver, varenr)
            entry.update(metrics)
        
        results.append(entry)

    # Cleanup and final save
    driver.quit()
    stop_event.set()
    save_thread.join()
    write_excel(results, args.output_xlsx)
    # Note: Logging statement removed to avoid exposing confidential data

if __name__ == "__main__":
    main()
