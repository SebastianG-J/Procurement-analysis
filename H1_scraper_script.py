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

SEARCH_INPUT_XPATH = "/html/body/header/div/div[1]/div[1]/div/form/input"
DROPDOWN_FIRST_ITEM_XPATH = "/html/body/header/div/div[1]/div[1]/div/div/a[1]"
DEFAULT_EXCEL_FILE = "vareposter_antal_salg_køb_modregning_opdateret_med_købt.xlsx"
OUTPUT_XLSX = "results_simplified.xlsx"
COOKIE_ACCEPT_XPATH = '//*[@id="coiPage-1"]/div[2]/div[1]/button[3]'
DEFAULT_TABLE_XPATH = '//*[@id="addMultipleToCartForm"]/div/table'

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

def init_driver(headless: bool):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1280,800")
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(1)
    return driver

def load_product_numbers_from_excel(path: str):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Excel file not found: {path}")
    df = pd.read_excel(path, engine="openpyxl")
    if "Varenr." not in df.columns:
        raise KeyError(f"'Varenr.' column not found in Excel. Columns present: {list(df.columns)}")
    raw = df["Varenr."].dropna()
    cleaned = raw.astype(str).str.strip()
    cleaned = cleaned[cleaned != ""]
    unique = cleaned.drop_duplicates()
    return unique.tolist()

def accept_cookies_if_needed(driver):
    try:
        wait = WebDriverWait(driver, 5)
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, COOKIE_ACCEPT_XPATH)))
        btn.click()
        WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((By.XPATH, COOKIE_ACCEPT_XPATH)))
    except TimeoutException:
        pass
    except Exception:
        pass

def search_and_open_product(driver, search_url, product_number):
    try:
        try:
            search_input = WebDriverWait(driver, 4).until(
                EC.element_to_be_clickable((By.XPATH, SEARCH_INPUT_XPATH))
            )
        except TimeoutException:
            driver.get(search_url)
            search_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, SEARCH_INPUT_XPATH))
            )
        search_input.clear()
        search_input.send_keys(product_number)
        first_suggestion = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, DROPDOWN_FIRST_ITEM_XPATH))
        )
        first_suggestion.click()
        time.sleep(0.5)
        return True
    except Exception:
        return False

def _normalize_header_text(text: str) -> str:
    return " ".join(text.lower().split())

def build_column_index_map(driver, table_xpath):
    headers = driver.find_elements(By.XPATH, f"{table_xpath}//thead//th")
    if not headers:
        headers = driver.find_elements(
            By.XPATH,
            f"{table_xpath}//tbody/tr[1]/th | {table_xpath}//tbody/tr[1]/td"
        )
    col_names = []
    for h in headers:
        txt = h.text.strip()
        if txt:
            col_names.append(_normalize_header_text(txt))
        else:
            col_names.append("")
    return {name: idx for idx, name in enumerate(col_names) if name}

def extract_table_metrics(driver, product_number, table_xpath=DEFAULT_TABLE_XPATH):
    result = {"meter_pr_rulle": "", "basisenhed": ""}
    col_map = build_column_index_map(driver, table_xpath)
    desired = {
        "meter pr. rulle": "meter_pr_rulle",
        "basisenhed": "basisenhed",
    }
    row = None
    try:
        row = driver.find_element(
            By.XPATH,
            f"{table_xpath}//tbody/tr[.//td[normalize-space()='{product_number}']]"
        )
    except Exception:
        try:
            row = driver.find_element(
                By.XPATH,
                f"{table_xpath}//tbody/tr[.//td[contains(normalize-space(.), '{product_number}')]]"
            )
        except Exception:
            return result  # not found
    cells = row.find_elements(By.XPATH, ".//td")
    for header_text, field_key in desired.items():
        idx = col_map.get(header_text)
        if idx is not None and idx < len(cells):
            result[field_key] = cells[idx].text.strip()
        else:
            # fallback based on known positions
            if field_key == "meter_pr_rulle":
                fallback_idx = 4  # td[5]
            elif field_key == "basisenhed":
                fallback_idx = 6  # td[7]
            else:
                fallback_idx = None
            if fallback_idx is not None and fallback_idx < len(cells):
                result[field_key] = cells[fallback_idx].text.strip()
    # Clean meter_pr_rulle: only keep if numerical
    if not result["meter_pr_rulle"].replace(",", ".").replace(".", "", 1).isdigit():
        result["meter_pr_rulle"] = ""
    # Clean basisenhed: only keep if "MTR" or "Mtr."
    if result["basisenhed"] not in ["MTR", "Mtr."]:
        result["basisenhed"] = ""
    return result

def write_outputs(results, xlsx_path):
    if not results:
        logging.warning("No results to write.")
        return
    try:
        df_out = pd.DataFrame(results)
        df_out.to_excel(xlsx_path, index=False)
        logging.info(f"Wrote Excel to {xlsx_path}")
    except Exception as e:
        logging.warning(f"Failed to write Excel output: {e}")

def periodic_save(results, xlsx_path, stop_event):
    while not stop_event.is_set():
        time.sleep(60)  # wait one minute
        write_outputs(results, xlsx_path)
        logging.info("Periodic save completed.")

def main():
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

    try:
        product_list = load_product_numbers_from_excel(args.excel_file)
        logging.info(f"Loaded {len(product_list)} unique product numbers from Excel.")
    except Exception as e:
        logging.error(f"Failed to load product numbers: {e}")
        return

    driver = init_driver(headless=args.headless)
    driver.get(args.search_url)
    accept_cookies_if_needed(driver)
    results = []

    # Start periodic save thread
    stop_event = threading.Event()
    save_thread = threading.Thread(target=periodic_save, args=(results, args.output_xlsx, stop_event))
    save_thread.daemon = True
    save_thread.start()

    for product_number in product_list:
        logging.info(f"Processing product number: {product_number}")
        entry = {
            "Varenr.": product_number,
            "meter_pr_rulle": "",
            "basisenhed": "",
        }
        success = search_and_open_product(driver, args.search_url, product_number)
        if success:
            metrics = extract_table_metrics(driver, product_number, DEFAULT_TABLE_XPATH)
            entry["meter_pr_rulle"] = metrics.get("meter_pr_rulle", "")
            entry["basisenhed"] = metrics.get("basisenhed", "")
        results.append(entry)
        delay = random.uniform(args.min_delay, args.max_delay)
        time.sleep(delay)

    driver.quit()
    stop_event.set()  # signal the save thread to stop
    save_thread.join()
    write_outputs(results, args.output_xlsx)
    logging.info("Finished all products.")

if __name__ == "__main__":
    main()