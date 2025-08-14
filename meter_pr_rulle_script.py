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

# ===== XPATHS =====
SEARCH_INPUT_XPATH = "/html/body/header/div/div[1]/div[1]/div/form/input"
DROPDOWN_FIRST_ITEM_XPATH = "/html/body/header/div/div[1]/div[1]/div/div/a[1]"
COOKIE_ACCEPT_XPATH = '//*[@id="coiPage-1"]/div[2]/div[1]/button[3]'
DEFAULT_TABLE_XPATH = '//*[@id="addMultipleToCartForm"]/div/table'

# ===== LOGGING =====
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# ===== DRIVER =====
def init_driver(headless: bool):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1280,800")
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(1)
    return driver

# ===== EXCEL LOADING =====
def load_varenr(path: str):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Excel file not found: {path}")
    df = pd.read_excel(path, engine="openpyxl")
    if "Varenr." not in df.columns:
        raise KeyError(f"'Varenr.' column not found. Found: {list(df.columns)}")
    values = df["Varenr."].dropna().astype(str).str.strip()
    values = values[values != ""].drop_duplicates()
    return set(values)

# ===== COOKIE =====
def accept_cookies(driver):
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, COOKIE_ACCEPT_XPATH))
        )
        btn.click()
        WebDriverWait(driver, 5).until(
            EC.invisibility_of_element_located((By.XPATH, COOKIE_ACCEPT_XPATH))
        )
    except TimeoutException:
        pass
    except Exception:
        pass

# ===== SEARCH =====
def search_and_open_product(driver, varenr):
    try:
        search_input = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, SEARCH_INPUT_XPATH))
        )
        search_input.clear()
        search_input.send_keys(varenr)
        first_item = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, DROPDOWN_FIRST_ITEM_XPATH))
        )
        first_item.click()
        time.sleep(0.5)
        return True
    except TimeoutException:
        return False

# ===== TABLE PARSING =====
def normalize_header(text: str) -> str:
    return " ".join(text.lower().split())

def build_col_map(driver, table_xpath):
    headers = driver.find_elements(By.XPATH, f"{table_xpath}//thead//th")
    if not headers:
        headers = driver.find_elements(
            By.XPATH,
            f"{table_xpath}//tbody/tr[1]/th | {table_xpath}//tbody/tr[1]/td"
        )
    col_map = {}
    for idx, h in enumerate(headers):
        txt = h.text.strip()
        if txt:
            col_map[normalize_header(txt)] = idx
    return col_map

def extract_metrics(driver, varenr, table_xpath=DEFAULT_TABLE_XPATH):
    result = {"meter_pr_rulle": "", "basisenhed": ""}
    col_map = build_col_map(driver, table_xpath)
    desired = {
        "meter pr. rulle": "meter_pr_rulle",
        "basisenhed": "basisenhed",
    }

    try:
        row = driver.find_element(
            By.XPATH,
            f"{table_xpath}//tbody/tr[.//td[normalize-space()='{varenr}']]"
        )
    except:
        try:
            row = driver.find_element(
                By.XPATH,
                f"{table_xpath}//tbody/tr[.//td[contains(normalize-space(.), '{varenr}')]]"
            )
        except:
            return result

    cells = row.find_elements(By.XPATH, ".//td")
    for header, key in desired.items():
        idx = col_map.get(header)
        if idx is not None and idx < len(cells):
            result[key] = cells[idx].text.strip()
        else:
            # fallback index
            if key == "meter_pr_rulle" and len(cells) > 4:
                result[key] = cells[4].text.strip()
            elif key == "basisenhed" and len(cells) > 6:
                result[key] = cells[6].text.strip()

    # cleanup
    if not result["meter_pr_rulle"].replace(",", ".").replace(".", "", 1).isdigit():
        result["meter_pr_rulle"] = ""
    if result["basisenhed"] not in ["MTR", "Mtr."]:
        result["basisenhed"] = ""
    return result

# ===== SAVE =====
def write_excel(results, path):
    if results:
        pd.DataFrame(results).to_excel(path, index=False)
        logging.info(f"Saved {len(results)} rows to {path}")

# ===== PERIODIC SAVE =====
def periodic_save(results, path, stop_event):
    while not stop_event.is_set():
        time.sleep(60)
        write_excel(results, path)
        logging.info("Periodic save complete.")

# ===== MAIN =====
def main():
    parser = argparse.ArgumentParser(description="Scrape new Varenr. from Alfotech.")
    parser.add_argument("--search-url", default="https://www.alfotech.dk/", help="Search page URL")
    parser.add_argument("--full-list", default="vareposter_antal_salg_køb_modregning_opdateret_med_købt.xlsx", help="Full product list")
    parser.add_argument("--exclude-xlsx", default="results_merged.xlsx", help="Already scraped list")
    parser.add_argument("--output-xlsx", default="results_new.xlsx", help="New output file")
    parser.add_argument("--headless", action="store_true")
    args = parser.parse_args()

    # Load Excel data
    full_list = load_varenr(args.full_list)
    exclude_list = load_varenr(args.exclude_xlsx)
    to_scrape = sorted(full_list - exclude_list)

    # Filter out any Varenr. that begins with "25"
    to_scrape = [v for v in to_scrape if not str(v).startswith("25")]

    logging.info(f"Found {len(full_list)} total, {len(exclude_list)} excluded, {len(to_scrape)} to scrape after filtering.")

    driver = init_driver(headless=args.headless)
    driver.get(args.search_url)
    accept_cookies(driver)

    results = []

    # Start periodic save thread
    stop_event = threading.Event()
    save_thread = threading.Thread(target=periodic_save, args=(results, args.output_xlsx, stop_event))
    save_thread.daemon = True
    save_thread.start()

    for varenr in to_scrape:
        logging.info(f"Scraping {varenr}...")
        entry = {"Varenr.": varenr, "meter_pr_rulle": "", "basisenhed": ""}
        if search_and_open_product(driver, varenr):
            metrics = extract_metrics(driver, varenr)
            entry.update(metrics)
        results.append(entry)

    # Cleanup
    driver.quit()
    stop_event.set()
    save_thread.join()
    write_excel(results, args.output_xlsx)
    logging.info("Done.")

if __name__ == "__main__":
    main()
