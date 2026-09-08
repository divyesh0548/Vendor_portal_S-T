from __future__ import annotations

import importlib.util
import logging
import re
import sys
import time
from pathlib import Path
from typing import Any

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select


def _load_gst_main_module():
    module_name = "gst_extraction_process_main"
    existing = sys.modules.get(module_name)
    if existing is not None:
        return existing

    main_path = Path(__file__).resolve().with_name("main.py")
    spec = importlib.util.spec_from_file_location(module_name, str(main_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load GST extractor main module from {main_path}")

    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)  # type: ignore[arg-type]
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    return module


gst_main = _load_gst_main_module()


WAIT_TIMEOUT = int(getattr(gst_main, "WAIT_TIMEOUT", 25) or 25)
logger = logging.getLogger(__name__)
NO_RECORD_XPATH = (
    "//div[@data-ng-if='noRecord']//p"
    "[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'no records found')]"
)
FILING_TABLE_SIGNAL_XPATH = (
    "//div[@class='table-responsive']//h4"
    "[contains(text(),'GSTR3B') or contains(text(),'GSTR-1/IFF') or contains(text(),'GSTR-1')]"
)


def normalize_text(value: object) -> str:
    return gst_main.normalize_text(value)


def _normalize_gst_status(value: Any) -> str:
    return normalize_text(value).lower()


def _extract_table_from_element(table) -> tuple[list[str], list[list[str]]]:
    headers: list[str] = []
    rows_data: list[list[str]] = []

    try:
        header_row = table.find_element(By.XPATH, ".//thead//tr[@class='ng-table-sort-header']")
        header_cells = header_row.find_elements(By.XPATH, ".//th//span[@class='sort-indicator']")
        headers = [normalize_text(cell.text) for cell in header_cells if normalize_text(cell.text)]
    except Exception:
        headers = []

    if not headers:
        try:
            first_data_row = table.find_element(By.XPATH, ".//tbody//tr[1]")
            header_cells = first_data_row.find_elements(By.XPATH, ".//td[@data-title-text]")
            headers = [
                normalize_text(cell.get_attribute("data-title-text"))
                for cell in header_cells
                if normalize_text(cell.get_attribute("data-title-text"))
            ]
        except Exception:
            headers = []

    try:
        tbody = table.find_element(By.TAG_NAME, "tbody")
        data_rows = tbody.find_elements(By.XPATH, ".//tr[td]")
        for row in data_rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            row_data = [normalize_text(cell.text) for cell in cells]
            if any(row_data):
                rows_data.append(row_data)
    except Exception:
        rows_data = []

    return headers, rows_data


def _extract_trade_name(driver: webdriver.Chrome) -> str:
    try:
        trade_div = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'col-sm-4') and contains(@class,'col-xs-12') and @data-ng-if='tradeFlag']",
        )
        p_tags = trade_div.find_elements(By.TAG_NAME, "p")
        if len(p_tags) >= 2:
            trade_name = normalize_text(p_tags[1].text)
            if trade_name and trade_name.lower() != "trade name":
                return trade_name
    except Exception:
        pass
    return ""


def _click_show_filing_table(driver: webdriver.Chrome, wait: WebDriverWait) -> bool:
    button_xpaths = [
        "//button[@id='filingTable']",
        "//button[@data-ng-click=\"getFinYearDropdown('prelogin')\"]",
    ]

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)

    for xpath in button_xpaths:
        try:
            button = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", button)
            time.sleep(0.5)
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            driver.execute_script("arguments[0].click();", button)
            time.sleep(2)
            return True
        except Exception:
            continue
    return False


def _click_filing_search(driver: webdriver.Chrome, wait: WebDriverWait) -> bool:
    xpath = "//button[@data-ng-click=\"getFilingData(finyr, 'prelogin')\"]"
    try:
        button = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
        time.sleep(0.5)
        button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].click();", button)
        time.sleep(0.3)
        return True
    except Exception:
        return False


def _normalize_financial_year(value: Any) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    match = re.search(r"(20\d{2})\D*(\d{2,4})", text)
    if not match:
        return text.lower()
    start_year = int(match.group(1))
    end_raw = str(match.group(2))
    if len(end_raw) == 4:
        end_year = int(end_raw) % 100
    else:
        end_year = int(end_raw)
    return f"{start_year}-{end_year:02d}"


def _select_financial_year(driver: webdriver.Chrome, wait: WebDriverWait, financial_year: Any) -> bool:
    target = _normalize_financial_year(financial_year)
    if not target:
        return True

    select_xpaths = [
        "//select[@data-ng-model='finyr']",
        "//select[@ng-model='finyr']",
        "//select[@id='finyear']",
        "//select[@name='finyear']",
        "//select[contains(@data-ng-options,'finyr')]",
    ]

    for xpath in select_xpaths:
        try:
            select_el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            wait.until(EC.visibility_of(select_el))
            select_obj = Select(select_el)
        except Exception:
            continue

        for option in select_obj.options:
            option_text = _normalize_financial_year(option.text)
            option_value = _normalize_financial_year(option.get_attribute("value"))
            if target in {option_text, option_value}:
                try:
                    select_obj.select_by_visible_text(option.text)
                except Exception:
                    try:
                        select_obj.select_by_value(option.get_attribute("value"))
                    except Exception:
                        continue
                time.sleep(0.4)
                return True

        try:
            start_target, end_target = target.split("-", 1)
        except ValueError:
            start_target, end_target = "", ""
        for option in select_obj.options:
            option_text_raw = normalize_text(option.text)
            option_value_raw = normalize_text(option.get_attribute("value"))
            candidate_strings = [option_text_raw, option_value_raw]
            if start_target and end_target and any(start_target in c and end_target in c for c in candidate_strings):
                try:
                    select_obj.select_by_visible_text(option.text)
                except Exception:
                    try:
                        select_obj.select_by_value(option.get_attribute("value"))
                    except Exception:
                        continue
                time.sleep(0.4)
                return True

    return False


def _extract_table_by_heading(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    heading_xpath: str,
    table_xpath: str,
    wait_timeout: int = WAIT_TIMEOUT,
) -> dict[str, Any] | None:
    try:
        local_wait = WebDriverWait(driver, max(1, int(wait_timeout)))
        heading = local_wait.until(EC.presence_of_element_located((By.XPATH, heading_xpath)))
        title = normalize_text(heading.text)
        local_wait.until(EC.presence_of_element_located((By.XPATH, table_xpath)))
        table = driver.find_element(By.XPATH, table_xpath)
        headers, rows = _extract_table_from_element(table)
        if not headers and not rows:
            return None
        return {
            "title": title,
            "headers": headers,
            "rows": rows,
        }
    except Exception:
        return None


def _is_no_record_visible(driver: webdriver.Chrome) -> bool:
    try:
        elements = driver.find_elements(By.XPATH, NO_RECORD_XPATH)
    except Exception:
        return False

    for element in elements:
        try:
            if element.is_displayed():
                return True
        except Exception:
            continue
    return False


def _wait_for_filing_result_state(driver: webdriver.Chrome, timeout: int = 8) -> str:
    try:
        WebDriverWait(driver, max(1, int(timeout))).until(
            lambda drv: _is_no_record_visible(drv) or bool(drv.find_elements(By.XPATH, FILING_TABLE_SIGNAL_XPATH))
        )
    except Exception:
        return "timeout"

    if _is_no_record_visible(driver):
        return "no_record"

    try:
        if driver.find_elements(By.XPATH, FILING_TABLE_SIGNAL_XPATH):
            return "table"
    except Exception:
        pass
    return "timeout"


def _extract_filing_tables(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    financial_year: Any = None,
) -> tuple[list[dict[str, Any]], bool]:
    tables: list[dict[str, Any]] = []
    no_record = False
    if not _click_show_filing_table(driver, wait):
        logger.warning("Unable to open GST filing table section.")
        return tables, no_record

    if financial_year and not _select_financial_year(driver, wait, financial_year):
        logger.warning("Unable to select GST filing financial year: %s", financial_year)

    if not _click_filing_search(driver, wait):
        return tables, _is_no_record_visible(driver)

    result_state = _wait_for_filing_result_state(driver, timeout=8)
    if result_state == "no_record":
        return tables, True

    table_specs = [
        {
            "heading_xpath": "//div[@class='table-responsive']//h4[contains(text(),'GSTR3B')]",
            "table_xpath": "//div[@class='table-responsive']//h4[contains(text(),'GSTR3B')]/ancestor::div[@class='table-responsive'][1]//table[1]",
        },
        {
            "heading_xpath": "//div[@class='table-responsive']//h4[contains(text(),'GSTR-1/IFF') or contains(text(),'GSTR-1')]",
            "table_xpath": "//div[@class='table-responsive']//h4[contains(text(),'GSTR-1/IFF') or contains(text(),'GSTR-1')]/ancestor::div[@class='table-responsive'][1]//table[1]",
        },
    ]

    table_wait_timeout = 6 if result_state == "table" else 3
    for spec in table_specs:
        if _is_no_record_visible(driver):
            no_record = True
            break
        table_data = _extract_table_by_heading(
            driver,
            wait,
            spec["heading_xpath"],
            spec["table_xpath"],
            wait_timeout=table_wait_timeout,
        )
        if table_data:
            tables.append(table_data)

    if not tables and _is_no_record_visible(driver):
        no_record = True

    return tables, no_record


def _extract_one_gst_filing(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    gst: str,
    financial_year: Any = None,
) -> dict[str, Any]:
    gst_summary = gst_main._extract_one_gst(driver, wait, gst)
    result: dict[str, Any] = {
        "GST Number": gst,
        "Trade Name": "",
        "Taxpayer Type": "",
        "GSTIN / UIN  Status": "Error",
        "Filing Message": "",
        "Filing Tables": [],
    }

    if isinstance(gst_summary, dict):
        result.update(
            {
                "GST Number": str(gst_summary.get("GST Number") or gst).strip() or gst,
                "Trade Name": str(gst_summary.get("Trade Name") or "").strip(),
                "Taxpayer Type": str(gst_summary.get("Taxpayer Type") or "").strip(),
                "GSTIN / UIN  Status": str(gst_summary.get("GSTIN / UIN  Status") or "").strip() or "Error",
            }
        )

    status_key = _normalize_gst_status(result.get("GSTIN / UIN  Status"))
    if status_key not in {"valid", "active"}:
        return result

    if not result["Trade Name"]:
        result["Trade Name"] = _extract_trade_name(driver)

    filing_tables, no_record = _extract_filing_tables(driver, wait, financial_year=financial_year)
    result["Filing Tables"] = filing_tables
    if no_record and not filing_tables:
        result["Filing Message"] = "Vendor GST Not File Return For This New Year"
    return result


def main(gst_list=None, save_to_excel=False, financial_year=None):
    """
    GST filing-table extraction entrypoint used by backend.

    main([gst_number], save_to_excel=False) -> list[dict]
    """
    if save_to_excel:
        logger.warning("save_to_excel=True ignored: Excel flow removed from this module.")

    if gst_list is None:
        logger.error("No GST numbers provided.")
        return []

    gst_values = [gst_main.normalize_gst_number(str(item)) for item in gst_list if str(item).strip()]
    if not gst_values:
        logger.error("No GST numbers provided.")
        return []

    run_results: list[dict[str, Any]] = []
    driver = webdriver.Chrome()
    driver.maximize_window()
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    try:
        for gst in gst_values:
            logger.info("Processing GST filing tables for %s", gst)
            run_results.append(_extract_one_gst_filing(driver, wait, gst, financial_year=financial_year))
            time.sleep(0.2)
    finally:
        logger.info("GST filing extraction complete (closing browser).")
        try:
            driver.quit()
        except Exception:
            pass

    return run_results
