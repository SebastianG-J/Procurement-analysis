"""
H1 Product Scraper - Documented Version
=======================================

This script is the main web scraper for collecting product specifications from supplier websites.
It automates the process of searching for products and extracting specific metrics like
meter per roll and base unit information.

Purpose:
- Automate product data collection from supplier websites
- Extract product specifications (meter per roll, base units)
- Handle large product lists efficiently with periodic saving
- Provide robust error handling and retry mechanisms

Key Features:
- Automated web scraping using Selenium WebDriver
- Periodic data saving to prevent data loss
- Configurable delays to respect website rate limits
- Intelligent product search and data extraction
- Comprehensive error handling and logging

Author: [Your Name]
Date: [Date]
"""

import argparse
import logging
import os
import time
import random
import threading
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ===== XPATH SELECTORS =====
# These XPath selectors target specific elements on the supplier website
SEARCH_INPUT_XPATH = "/html/body/header/div/div[1]/div[1]/div/form/input"
DROPDOWN_FIRST_ITEM_XPATH = "/html/body/header/div/div[1]/div[1]/div/div/a[1]"
DEFAULT_EXCEL_FILE = "vareposter_antal_salg_køb_modregning_opdateret_med_købt.xlsx"
OUTPUT_XLSX = "results_simplified.xlsx"
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
def load_product_numbers_from_excel(path: str):
    """
    Load product numbers (Varenr.) from an Excel file.
    
    Args:
        path (str): Path to the Excel file containing product numbers
    
    Returns:
        list: List of unique product numbers
    
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
        raise KeyError(f"'Varenr.' column not found in Excel. Columns present: {list(df.columns)}")
    
    # Clean and process the product numbers
    raw = df["Varenr."].dropna()
    cleaned = raw.astype(str).str.strip()
    cleaned = cleaned[cleaned != ""]
    unique = cleaned.drop_duplicates()
    return unique.tolist()

# ===== COOKIE HANDLING =====
def accept_cookies_if_needed(driver):
    """
    Accept cookies on the website to avoid popup interference.
    
    Args:
        driver: Chrome web driver instance
    """
    try:
        # Wait for cookie accept button to be clickable
        wait = WebDriverWait(driver, 5)
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, COOKIE_ACCEPT_XPATH)))
        btn.click()
        
        # Wait for cookie popup to disappear
        WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((By.XPATH, COOKIE_ACCEPT_XPATH)))
    except TimeoutException:
        # Cookie popup might not appear, continue silently
        pass
    except Exception:
        # Any other error, continue silently
        pass

# ===== PRODUCT SEARCH =====
def search_and_open_product(driver, search_url, product_number):
    """
    Search for a product on the website and open its details page.
    
    Args:
        driver: Chrome web driver instance
        search_url (str): URL of the search page
        product_number (str): Product number to search for
    
    Returns:
        bool: True if product was found and opened, False otherwise
    """
    try:
        try:
            # Wait for search input to be available and clickable
            search_input = WebDriverWait(driver, 4).until(
                EC.element_to_be_clickable((By.XPATH, SEARCH_INPUT_XPATH))
            )
        except TimeoutException:
            # If search input not found, navigate to search URL and try again
            driver.get(search_url)
            search_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, SEARCH_INPUT_XPATH))
            )
        
        # Clear any existing text and enter the product number
        search_input.clear()
        search_input.send_keys(product_number)
        
        # Wait for and click the first search result
        first_suggestion = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, DROPDOWN_FIRST_ITEM_XPATH))
        )
        first_suggestion.click()
        
        # Brief pause to let the page load
        time.sleep(0.5)
        return True
        
    except Exception:
        # Product not found or search failed
        return False

# ===== TABLE PARSING =====
def _normalize_header_text(text: str) -> str:
    """
    Normalize table header text for consistent matching.
    
    Args:
        text (str): Raw header text
    
    Returns:
        str: Normalized header text (lowercase, single spaces)
    """
    return " ".join(text.lower().split())

def build_column_index_map(driver, table_xpath):
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
    col_names = []
    for h in headers:
        txt = h.text.strip()
        if txt:
            col_names.append(_normalize_header_text(txt))
        else:
            col_names.append("")
    
    return {name: idx for idx, name in enumerate(col_names) if name}

def extract_table_metrics(driver, product_number, table_xpath=DEFAULT_TABLE_XPATH):
    """
    Extract meter per roll and base unit information from the product table.
    
    Args:
        driver: Chrome web driver instance
        product_number (str): Product number to extract data for
        table_xpath (str): XPath to the table containing product data
    
    Returns:
        dict: Dictionary containing 'meter_pr_rulle' and 'basisenhed' values
    """
    # Initialize result with empty values
    result = {"meter_pr_rulle": "", "basisenhed": ""}
    
    # Build column mapping for the table
    col_map = build_column_index_map(driver, table_xpath)
    
    # Define which columns we want to extract
    desired = {
        "meter pr. rulle": "meter_pr_rulle",
        "basisenhed": "basisenhed",
    }
    
    # Try to find the row containing the product number
    row = None
    try:
        # Exact match first
        row = driver.find_element(
            By.XPATH,
            f"{table_xpath}//tbody/tr[.//td[normalize-space()='{product_number}']]"
        )
    except Exception:
        try:
            # Partial match if exact match fails
            row = driver.find_element(
                By.XPATH,
                f"{table_xpath}//tbody/tr[.//td[contains(normalize-space(.), '{product_number}')]]"
            )
        except Exception:
            return result  # Product not found
    
    # Extract data from the found row
    cells = row.find_elements(By.XPATH, ".//td")
    
    # Extract values for each desired column
    for header_text, field_key in desired.items():
        idx = col_map.get(header_text)
        if idx is not None and idx < len(cells):
            result[field_key] = cells[idx].text.strip()
        else:
            # Fallback to hardcoded indices if column mapping fails
            if field_key == "meter_pr_rulle":
                fallback_idx = 4  # td[5]
            elif field_key == "basisenhed":
                fallback_idx = 6  # td[7]
            else:
                fallback_idx = None
            
            if fallback_idx is not None and fallback_idx < len(cells):
                result[field_key] = cells[fallback_idx].text.strip()
    
    # Clean up and validate the extracted data
    # Check if meter_pr_rulle is a valid number
    if not result["meter_pr_rulle"].replace(",", ".").replace(".", "", 1).isdigit():
        result["meter_pr_rulle"] = ""
    
    # Check if basisenhed is a valid unit
    if result["basisenhed"] not in ["MTR", "Mtr."]:
        result["basisenhed"] = ""
    
    return result

# ===== DATA SAVING =====
def write_outputs(results, xlsx_path):
    """
    Save the scraped results to an Excel file.
    
    Args:
        results (list): List of dictionaries containing scraped data
        xlsx_path (str): Path where the Excel file should be saved
    """
    if not results:
        logging.warning("No results to write.")
        return
    
    try:
        df_out = pd.DataFrame(results)
        df_out.to_excel(xlsx_path, index=False)
        # Note: Logging statement removed to avoid exposing confidential data
    except Exception as e:
        logging.warning(f"Failed to write Excel output: {e}")

# ===== PERIODIC SAVING =====
def periodic_save(results, xlsx_path, stop_event):
    """
    Periodically save results to prevent data loss during long scraping sessions.
    
    Args:
        results (list): List of scraped results (shared between threads)
        xlsx_path (str): Path where to save the data
        stop_event (threading.Event): Event to signal when to stop saving
    """
    while not stop_event.is_set():
        time.sleep(60)  # Save every 60 seconds
        write_outputs(results, xlsx_path)
        # Note: Logging statement removed to avoid exposing confidential data

# ===== MAIN EXECUTION =====
def main():
    """
    Main function that orchestrates the entire scraping process.
    """
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description="Selenium scraper for product metrics.")
    parser.add_argument("--search-url", required=True, help="URL with the static search bar.")
    parser.add_argument(
        "--excel-file",
        default=DEFAULT_EXCEL_FILE,
        help=f"Excel file containing product numbers in column 'Varenr.' (default: {DEFAULT_EXCEL_FILE})",
    )
    parser.add_argument("--output-xlsx", default=OUTPUT_XLSX, help="Path for Excel output.")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode (no UI).")
    parser.add_argument("--min-delay", type=float, default=0.5, help="Minimum delay between products.")
    parser.add_argument("--max-delay", type=float, default=1.0, help="Maximum delay between products.")

    args = parser.parse_args()

    # Load product list from Excel file
    try:
        product_list = load_product_numbers_from_excel(args.excel_file)
        # Note: Logging statement removed to avoid exposing confidential data
    except Exception as e:
        logging.error(f"Failed to load product numbers: {e}")
        return

    # Initialize web driver and navigate to search page
    driver = init_driver(headless=args.headless)
    driver.get(args.search_url)
    accept_cookies_if_needed(driver)
    results = []

    # Start periodic save thread
    stop_event = threading.Event()
    save_thread = threading.Thread(target=periodic_save, args=(results, args.output_xlsx, stop_event))
    save_thread.daemon = True
    save_thread.start()

    # Scrape each product
    for product_number in product_list:
        # Note: Logging statement removed to avoid exposing confidential data
        
        # Initialize entry for this product
        entry = {
            "Varenr.": product_number,
            "meter_pr_rulle": "",
            "basisenhed": "",
        }
        
        # Try to scrape the product data
        success = search_and_open_product(driver, args.search_url, product_number)
        if success:
            metrics = extract_table_metrics(driver, product_number, DEFAULT_TABLE_XPATH)
            entry["meter_pr_rulle"] = metrics.get("meter_pr_rulle", "")
            entry["basisenhed"] = metrics.get("basisenhed", "")
        
        results.append(entry)
        
        # Random delay to respect website rate limits
        delay = random.uniform(args.min_delay, args.max_delay)
        time.sleep(delay)

    # Cleanup and final save
    driver.quit()
    stop_event.set()  # Signal the save thread to stop
    save_thread.join()
    write_outputs(results, args.output_xlsx)
    # Note: Logging statement removed to avoid exposing confidential data

if __name__ == "__main__":
    main()
