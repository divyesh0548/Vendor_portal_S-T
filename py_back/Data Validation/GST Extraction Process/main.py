import base64
import logging
import os
import re
import time
from io import BytesIO

import numpy as np
from PIL import Image
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# YOUR OCR MODULE - do NOT change
from ocr_module import extract_text, special_character_remover


# No Excel input/output in this module.
# Backend uses: main([gst_number], save_to_excel=False)

captcha_file = "captcha.png"
max_captcha_attempts = 8
WAIT_TIMEOUT = 25
GST_REGEX = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")

required_headers = [
    "GST Number",
    "Trade Name",
    "Taxpayer Type",
    "GSTIN / UIN  Status",
]

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def normalize_text(s: object) -> str:
    return " ".join(str(s or "").split()).strip()


def normalize_gst_number(gst: str) -> str:
    return normalize_text(gst).upper().replace(" ", "")


def header_key(label: str) -> str | None:
    lbl = normalize_text(label).lower()
    if "trade name" in lbl and "additional" not in lbl:
        return "Trade Name"
    if "taxpayer type" in lbl:
        return "Taxpayer Type"
    if "gstin" in lbl or "uin" in lbl or "status" in lbl:
        return "GSTIN / UIN  Status"
    return None


def extract_from_div_block(block) -> tuple[str, str]:
    label = ""
    try:
        strong = block.find_element(By.TAG_NAME, "strong")
        label = normalize_text(strong.text)
    except Exception:
        label = ""

    try:
        ul = block.find_element(By.XPATH, ".//ul[contains(@class,'jurisdictList')]")
        lis = ul.find_elements(By.TAG_NAME, "li")
        values = [normalize_text(li.text) for li in lis if normalize_text(li.text)]
        return label, " | ".join(values)
    except Exception:
        pass

    try:
        ps = block.find_elements(By.TAG_NAME, "p")
        if len(ps) >= 2:
            return label, normalize_text(ps[1].text)
        if len(ps) == 1:
            ptxt = normalize_text(ps[0].text)
            if label and label in ptxt:
                return label, normalize_text(ptxt.replace(label, ""))
            return label, ptxt
    except Exception:
        pass

    return label, ""


def click_refresh_button(driver_ref: webdriver.Chrome) -> bool:
    try:
        btn = driver_ref.find_element(By.XPATH, "//button[@ng-click='refreshCaptcha()']")
        driver_ref.execute_script("arguments[0].click();", btn)
        time.sleep(0.7)
        return True
    except Exception:
        return False


def _extract_one_gst(driver: webdriver.Chrome, wait: WebDriverWait, gst: str) -> dict[str, str]:
    data_map = {h: "" for h in required_headers}
    data_map["GST Number"] = gst

    if not GST_REGEX.fullmatch(gst):
        logger.warning("Skipping %s due to invalid GST format.", gst)
        data_map["GSTIN / UIN  Status"] = "Invalid"
        return data_map

    driver.get("https://services.gst.gov.in/services/searchtp")
    try:
        wait.until(EC.element_to_be_clickable((By.ID, "for_gstin")))
    except TimeoutException:
        logger.warning("GST input field not available - skipping %s", gst)
        data_map["GSTIN / UIN  Status"] = "Error"
        return data_map

    gst_input = driver.find_element(By.ID, "for_gstin")
    gst_input.clear()
    gst_input.send_keys(gst)

    if os.path.exists(captcha_file):
        try:
            os.remove(captcha_file)
        except Exception:
            pass

    passed = False
    invalid_gst = False

    for attempt in range(1, max_captcha_attempts + 1):
        logger.info("Attempt %d for GST %s", attempt, gst)
        try:
            wait.until(EC.presence_of_element_located((By.ID, "imgCaptcha")))
            wait.until(
                lambda drv: drv.execute_script(
                    "var img=document.getElementById('imgCaptcha'); return img && img.complete && img.naturalWidth>0"
                )
            )
        except Exception:
            if not click_refresh_button(driver):
                logger.warning("Refresh button not found - retrying...")
            time.sleep(0.7)
            continue

        captcha_base64 = ""
        try:
            captcha_base64 = driver.execute_script(
                "var img=document.getElementById('imgCaptcha'); if(!img) return ''; var c=document.createElement('canvas'); c.width=img.naturalWidth||img.width; c.height=img.naturalHeight||img.height; c.getContext('2d').drawImage(img,0,0); return c.toDataURL('image/png').split(',')[1];"
            )
        except Exception as exc:
            logger.debug("Canvas extraction issue: %s", exc)

        if not captcha_base64:
            if not click_refresh_button(driver):
                logger.warning("Refresh button not found - retrying...")
            time.sleep(0.7)
            continue

        try:
            img = Image.open(BytesIO(base64.b64decode(captcha_base64)))
            img.save(captcha_file)
        except Exception as exc:
            logger.error("Failed to save captcha image: %s", exc)
            if not click_refresh_button(driver):
                logger.warning("Refresh button not found - retrying...")
            time.sleep(0.7)
            continue

        try:
            np_img = np.array(img)
            raw = extract_text(np_img)
            final = special_character_remover(" ".join(raw) if isinstance(raw, (list, tuple)) else str(raw))
            final = normalize_text(final)
        except Exception as exc:
            logger.error("OCR failed: %s", exc)
            final = ""

        if not final:
            if not click_refresh_button(driver):
                logger.warning("Refresh button not found - retrying...")
            time.sleep(0.7)
            continue

        try:
            cap_input = wait.until(EC.element_to_be_clickable((By.ID, "fo-captcha")))
            cap_input.clear()
            cap_input.send_keys(final)
            search_btn = wait.until(EC.element_to_be_clickable((By.ID, "lotsearch")))
            search_btn.click()
        except Exception as exc:
            logger.error("Failed to submit captcha/search: %s", exc)
            if not click_refresh_button(driver):
                logger.warning("Refresh button not found - retrying...")
            time.sleep(0.7)
            continue

        try:
            wait_short = WebDriverWait(driver, 6)

            def either_condition(drv):
                try:
                    if drv.find_elements(By.XPATH, "//strong[contains(normalize-space(.),'Legal Name of Business')]"):
                        return "result"
                except Exception:
                    pass
                try:
                    elems = drv.find_elements(
                        By.XPATH,
                        "//span[@data-ng-bind='trans.ERR_GSTIN' or contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'the gstin/uin that you have entered is invalid')]",
                    )
                    for elem in elems:
                        if normalize_text(elem.text):
                            return "invalid_gst"
                except Exception:
                    pass
                try:
                    cap_elems = drv.find_elements(
                        By.XPATH,
                        "//span[@data-ng-if='searchtaxp.cap.$error.invalid_captcha' or contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'enter valid letters')]",
                    )
                    for elem in cap_elems:
                        if normalize_text(elem.text):
                            return "invalid_captcha"
                except Exception:
                    pass
                return False

            res = wait_short.until(either_condition)
        except TimeoutException:
            res = None
            try:
                if wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//strong[contains(normalize-space(.),'Legal Name of Business')]")
                    )
                ):
                    res = "result"
            except Exception:
                res = None

        if res == "invalid_gst":
            invalid_gst = True
            break
        if res == "invalid_captcha":
            click_refresh_button(driver)
            continue
        if res == "result":
            passed = True
            break

        try:
            driver.execute_script("if(typeof refreshCaptcha === 'function') refreshCaptcha();")
        except Exception:
            pass
        time.sleep(0.6)

    try:
        if os.path.exists(captcha_file):
            os.remove(captcha_file)
    except Exception:
        pass

    if invalid_gst:
        data_map["GSTIN / UIN  Status"] = "Invalid"
        return data_map
    if not passed:
        data_map["GSTIN / UIN  Status"] = "Error"
        return data_map

    try:
        blocks = driver.find_elements(By.XPATH, "//div[contains(@class,'col-sm-4') and contains(@class,'col-xs-12')]")
        for block in blocks:
            try:
                lbl, val = extract_from_div_block(block)
                if not lbl:
                    continue
                key = header_key(lbl)
                if not key:
                    continue
                if data_map.get(key):
                    if val and val not in data_map[key]:
                        data_map[key] = data_map[key] + " | " + val
                else:
                    data_map[key] = val
            except Exception:
                continue
    except Exception:
        logger.debug("No generic blocks found or extraction error")

    raw_status = normalize_text(data_map.get("GSTIN / UIN  Status", "")).lower()
    if raw_status in {"active", "valid"}:
        data_map["GSTIN / UIN  Status"] = "Valid"
    elif raw_status in {"invalid", "cancelled", "canceled"}:
        data_map["GSTIN / UIN  Status"] = "Invalid"
    else:
        data_map["GSTIN / UIN  Status"] = "Error"

    return data_map


def main(gst_list=None, save_to_excel=False):
    """
    GST extraction entrypoint used by backend.

    main([gst_number], save_to_excel=False) -> list[dict]
    """
    if save_to_excel:
        logger.warning("save_to_excel=True ignored: Excel flow removed from this module.")

    if gst_list is None:
        logger.error("No GST numbers provided.")
        return []

    gst_values = [normalize_gst_number(str(x)) for x in gst_list if str(x).strip()]
    if not gst_values:
        logger.error("No GST numbers provided.")
        return []

    run_results: list[dict[str, str]] = []
    driver = webdriver.Chrome()
    driver.maximize_window()
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    try:
        for gst in gst_values:
            logger.info("Processing GST: %s", gst)
            run_results.append(_extract_one_gst(driver, wait, gst))
            time.sleep(0.2)
    finally:
        logger.info("All done (closing browser).")
        try:
            driver.quit()
        except Exception:
            pass

    return run_results
