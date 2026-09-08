import re
import time
from io import BytesIO

from PIL import Image, ImageFilter, ImageOps
import pytesseract
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# Single Udyam lookup used by app/main.py.
# This keeps the app-facing run_single_udyam() API while using the more robust
# page flow and extraction logic from the external batch script.

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\Maulik Khunt\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
)

LANDING_URL = "https://udyamregistration.gov.in/Default.aspx"
VERIFY_URL = "https://udyamregistration.gov.in/Udyam_Verify.aspx"
PAGE_WAIT = 40
CAPTCHA_MAX_ATTEMPTS = 10
PER_NUMBER_REOPEN_LIMIT = 10
UDYAM_REGEX = re.compile(r"^UDYAM-[A-Z0-9]{2}-[0-9]{2}-[0-9]{7}$")
CAPTCHA_OCR_CONFIGS = (
    "--oem 3 --psm 7",
    "--oem 3 --psm 8",
    "--oem 3 --psm 13",
)
CAPTCHA_MIN_LENGTH = 4
MSME_CATEGORY_LABELS = {
    "micro": "Micro",
    "small": "Small",
    "medium": "Medium",
    "not an msme": "Not an MSME",
}
MSME_INDUSTRY_LABELS = {
    "manufacturing": "Manufacturing",
    "service": "Service",
    "services": "Service",
    "trading": "Trading",
    "trade": "Trading",
}

CAPTCHA_IMAGE_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_imgCaptcha"),
    (By.ID, "ctl00_ContentPlaceHolder1_captchaimg"),
]
CAPTCHA_INPUT_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_txtCaptcha"),
]
VERIFY_BUTTON_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_btnVerify"),
    (By.ID, "ctl00_ContentPlaceHolder1_btnValidate"),
]
REFRESH_CAPTCHA_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_ImgRefresh"),
    (By.ID, "ctl00_ContentPlaceHolder1_btnRefreshCaptcha"),
    (By.ID, "ctl00_ContentPlaceHolder1_btnRefresh"),
]
ENTERPRISE_NAME_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_lblEnterpriseName"),
    (By.XPATH, "//span[contains(@id,'EnterpriseName')]"),
]
UDYAM_INPUT_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_txtUdyamNo"),
]
VERIFY_LINK_LOCATORS = [
    (By.XPATH, "//a[contains(@href,'Udyam_Verify.aspx')]"),
]
STATUS_MESSAGE_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_lblMessage"),
    (By.ID, "ctl00_ContentPlaceHolder1_lblCaptcha"),
    (By.ID, "ctl00_ContentPlaceHolder1_divmgs"),
]
UDYAM_REQUIRED_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_RequiredFieldValidator50"),
]
UDYAM_FORMAT_ERROR_LOCATORS = [
    (By.ID, "ctl00_ContentPlaceHolder1_revUamNo"),
]
CANCELLATION_MESSAGE = (
    "The mentioned Udyam Registration Number has been cancelled by District Industries Centre or MSME-DFO."
)


def clean_text_keep_alnum(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(text or "")).strip()


def normalize_udyam_number(value: str) -> str:
    return str(value or "").strip().upper().replace(" ", "")


def normalize_msme_category(value: str) -> str:
    text = str(value or "").strip()
    lower = text.lower()
    for candidate, normalized in MSME_CATEGORY_LABELS.items():
        if candidate in lower:
            return normalized
    return ""


def normalize_msme_industry(value: str) -> str:
    text = str(value or "").strip()
    lower = text.lower()
    if "manufact" in lower:
        return "Manufacturing"
    for candidate, normalized in MSME_INDUSTRY_LABELS.items():
        if candidate in lower:
            return normalized
    return ""


def resize_for_ocr(img_pil: Image.Image, multiplier: int = 3) -> Image.Image:
    width = max(1, img_pil.width * multiplier)
    height = max(1, img_pil.height * multiplier)
    resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
    return img_pil.resize((width, height), resampling)


def build_captcha_variants(img_pil: Image.Image) -> list[Image.Image]:
    gray = ImageOps.autocontrast(img_pil.convert("L"))
    enlarged = resize_for_ocr(gray)
    threshold_dark = enlarged.point(lambda px: 0 if px < 150 else 255, mode="L")
    threshold_mid = enlarged.point(lambda px: 0 if px < 175 else 255, mode="L")
    sharpened = enlarged.filter(ImageFilter.SHARPEN)
    median = enlarged.filter(ImageFilter.MedianFilter(size=3))
    inverted = ImageOps.invert(threshold_mid.convert("L"))
    return [
        gray,
        enlarged,
        threshold_dark,
        threshold_mid,
        sharpened,
        median,
        inverted,
    ]


def score_captcha_candidate(text: str) -> tuple[int, int, int, int]:
    candidate = clean_text_keep_alnum(text)
    has_alpha = any(ch.isalpha() for ch in candidate)
    has_digit = any(ch.isdigit() for ch in candidate)
    preferred_length = 1 if 5 <= len(candidate) <= 8 else 0
    return (
        preferred_length,
        1 if has_alpha and has_digit else 0,
        min(len(candidate), 8),
        -abs(len(candidate) - 6),
    )


def ocr_read(img_pil: Image.Image) -> tuple[str, list[str]]:
    candidates: list[str] = []
    seen: set[str] = set()

    for variant in build_captcha_variants(img_pil):
        for config in CAPTCHA_OCR_CONFIGS:
            raw = pytesseract.image_to_string(variant, config=config)
            cleaned = clean_text_keep_alnum(raw)
            if not cleaned or cleaned in seen:
                continue
            seen.add(cleaned)
            candidates.append(cleaned)

    if not candidates:
        return "", []

    best_candidate = max(candidates, key=score_captcha_candidate)
    return best_candidate, candidates


def check_server_down(driver: webdriver.Chrome) -> bool:
    try:
        if driver.find_elements(By.ID, "main-frame-error"):
            return True
    except Exception:
        pass
    try:
        title = (driver.title or "").lower()
        if "can't be reached" in title or "this site can" in title and "reached" in title:
            return True
    except Exception:
        pass
    try:
        body = driver.find_element(By.TAG_NAME, "body").text.lower()
        signals = [
            "this site can",
            "err_timed_out",
            "took too long to respond",
            "server not found",
        ]
        if any(signal in body for signal in signals):
            return True
    except Exception:
        pass
    return False


def is_connection_refused_error(error: Exception) -> bool:
    text = str(error).lower()
    indicators = [
        "winerror 10061",
        "failed to establish a new connection",
        "max retries exceeded",
        "newconnectionerror",
        "connection refused",
        "connection aborted",
        "invalid session id",
        "disconnected: not connected to devtools",
    ]
    return any(indicator in text for indicator in indicators)


def start_chrome() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver


def stable_click(driver: webdriver.Chrome, element, *, allow_scroll: bool = False) -> None:
    if allow_scroll:
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                element,
            )
            time.sleep(0.2)
        except Exception:
            pass

    try:
        element.click()
        return
    except Exception:
        pass

    driver.execute_script("arguments[0].click();", element)


def js_click(driver: webdriver.Chrome, element) -> None:
    stable_click(driver, element)


def wait_for_first(
    driver: webdriver.Chrome,
    locators: list[tuple[str, str]],
    *,
    timeout: int = PAGE_WAIT,
    clickable: bool = False,
):
    for by, value in locators:
        try:
            condition = (
                EC.element_to_be_clickable((by, value))
                if clickable
                else EC.presence_of_element_located((by, value))
            )
            return WebDriverWait(driver, timeout).until(condition)
        except Exception:
            continue
    return None


def find_first(driver: webdriver.Chrome, locators: list[tuple[str, str]]):
    for by, value in locators:
        try:
            return driver.find_element(by, value)
        except Exception:
            continue
    return None


def get_text_safe(driver: webdriver.Chrome, element_id: str) -> str:
    try:
        return driver.find_element(By.ID, element_id).text.strip()
    except Exception:
        return ""


def body_text(driver: webdriver.Chrome) -> str:
    try:
        return str(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def body_contains_any(driver: webdriver.Chrome, phrases: list[str]) -> bool:
    text = body_text(driver).lower()
    return any(str(phrase or "").strip().lower() in text for phrase in phrases if phrase)


def read_status_message(driver: webdriver.Chrome) -> str:
    parts: list[str] = []
    for element in [find_first(driver, [locator]) for locator in STATUS_MESSAGE_LOCATORS]:
        if element is None:
            continue
        try:
            text = str(element.text or "").strip()
        except Exception:
            text = ""
        if text:
            parts.append(text)
    return " | ".join(parts)


def extract_labeled_row_value(driver: webdriver.Chrome, labels: tuple[str, ...]) -> str:
    normalized_labels = {str(label or "").strip().lower() for label in labels if str(label or "").strip()}
    if not normalized_labels:
        return ""

    try:
        rows = driver.find_elements(By.XPATH, "//tr")
    except Exception:
        return ""

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./th|./td")
        except Exception:
            continue
        if len(cells) < 2:
            continue

        texts: list[str] = []
        for cell in cells:
            try:
                texts.append(str(cell.text or "").strip())
            except Exception:
                texts.append("")

        for index, text in enumerate(texts[:-1]):
            lowered = text.lower()
            if lowered not in normalized_labels:
                continue
            next_text = texts[index + 1].strip()
            if next_text and next_text.lower() not in normalized_labels:
                return next_text
    return ""


def is_visible_with_text(driver: webdriver.Chrome, locators: list[tuple[str, str]], phrases: list[str]) -> bool:
    normalized_phrases = [str(phrase or "").strip().lower() for phrase in phrases if str(phrase or "").strip()]
    for locator in locators:
        element = find_first(driver, [locator])
        if element is None:
            continue
        try:
            if not element.is_displayed():
                continue
            text = str(element.text or "").strip().lower()
        except Exception:
            continue
        if not normalized_phrases:
            if text:
                return True
            continue
        if any(phrase in text for phrase in normalized_phrases):
            return True
    return False


def extract_classification_enterprise_type(driver: webdriver.Chrome) -> str:
    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        for table in tables:
            try:
                header_rows = table.find_elements(By.XPATH, ".//tr[th]")
                if not header_rows:
                    continue
                headers = [
                    str(th.text or "").strip()
                    for th in header_rows[0].find_elements(By.TAG_NAME, "th")
                ]
                if not headers:
                    continue

                header_texts = [header.lower() for header in headers]
                required_headers = (
                    "classification year",
                    "enterprise type",
                    "classification date",
                )
                if not all(any(required in header for header in header_texts) for required in required_headers):
                    continue

                enterprise_col_index = None
                for index, header in enumerate(header_texts):
                    if "enterprise type" in header:
                        enterprise_col_index = index
                        break
                if enterprise_col_index is None:
                    continue

                data_rows = table.find_elements(By.XPATH, ".//tr[td]")
                for row in data_rows:
                    if row.find_elements(By.TAG_NAME, "th"):
                        continue
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if enterprise_col_index >= len(cells):
                        continue
                    value = str(cells[enterprise_col_index].text or "").strip()
                    if normalize_msme_category(value):
                        return value
            except Exception:
                continue
    except Exception:
        pass
    return extract_labeled_row_value(driver, ("Enterprise Type",))


def extract_udyam_registration_date(driver: webdriver.Chrome) -> str:
    value = get_text_safe(driver, "ctl00_ContentPlaceHolder1_lblACKNOWLEDGEMENT")
    if value:
        return value
    return ""


def extract_major_activity(driver: webdriver.Chrome) -> str:
    for element_id in (
        "ctl00_ContentPlaceHolder1_lblServices",
        "ctl00_ContentPlaceHolder1_lblManufacturing",
        "ctl00_ContentPlaceHolder1_lbl_Activity",
    ):
        value = get_text_safe(driver, element_id)
        if value:
            return value
    return extract_labeled_row_value(driver, ("Major Activity",))


def refresh_captcha_if_available(driver: webdriver.Chrome) -> None:
    refresh_button = find_first(driver, REFRESH_CAPTCHA_LOCATORS)
    if refresh_button is None:
        return
    try:
        stable_click(driver, refresh_button)
        time.sleep(0.8)
    except Exception:
        pass


def open_verify_page(driver: webdriver.Chrome):
    driver.get(LANDING_URL)
    WebDriverWait(driver, PAGE_WAIT).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    if check_server_down(driver):
        raise RuntimeError("ServerDownOnLandingPage")

    verify_link = wait_for_first(driver, VERIFY_LINK_LOCATORS, timeout=PAGE_WAIT)
    if verify_link is None:
        raise RuntimeError("VerifyLinkNotFound")

    try:
        stable_click(driver, verify_link)
    except Exception:
        verify_link.click()

    input_element = wait_for_first(driver, UDYAM_INPUT_LOCATORS, timeout=PAGE_WAIT)
    if input_element is None:
        raise RuntimeError("UdyamInputNotFoundOnVerifyPage")
    return input_element


def run_single_udyam(udyam_number: str) -> dict[str, str]:
    """
    Returns vendor-update MSME fields for one Udyam number:
      - MSME Category
      - MSME Industry
      - Date of Udyam Registration
    """
    normalized_udyam = normalize_udyam_number(udyam_number)
    if not normalized_udyam or not UDYAM_REGEX.fullmatch(normalized_udyam):
        print(f"[MSME single] Invalid Udyam format: {normalized_udyam!r}")
        return {}

    driver = start_chrome()
    try:
        result: dict[str, str] = {}

        for attempt in range(1, PER_NUMBER_REOPEN_LIMIT + 1):
            try:
                print(f"[MSME single] Attempt {attempt}/{PER_NUMBER_REOPEN_LIMIT} for {normalized_udyam}")
                udyam_input = open_verify_page(driver)
                if check_server_down(driver):
                    raise RuntimeError("ServerDownOnVerifyPage")

                udyam_input.clear()
                udyam_input.send_keys(normalized_udyam)
                time.sleep(0.8)

                captcha_ok = False
                wrong_udyam_detected = False

                for captcha_attempt in range(1, CAPTCHA_MAX_ATTEMPTS + 1):
                    print(f"[MSME single] Captcha attempt {captcha_attempt}/{CAPTCHA_MAX_ATTEMPTS}")
                    if check_server_down(driver):
                        raise RuntimeError("ServerDownDuringCaptcha")

                    captcha_image = wait_for_first(driver, CAPTCHA_IMAGE_LOCATORS, timeout=20)
                    if captcha_image is None:
                        raise RuntimeError("CaptchaImageMissing")

                    captcha_png = captcha_image.screenshot_as_png
                    captcha_text, captcha_candidates = ocr_read(Image.open(BytesIO(captcha_png)))
                    print(f"[MSME single] OCR candidates: {captcha_candidates!r}")
                    print(f"[MSME single] OCR selected: {captcha_text!r}")
                    if len(captcha_text) < CAPTCHA_MIN_LENGTH:
                        refresh_captcha_if_available(driver)
                        time.sleep(0.8)
                        continue

                    captcha_input = find_first(driver, CAPTCHA_INPUT_LOCATORS)
                    if captcha_input is None:
                        raise RuntimeError("CaptchaInputMissing")
                    captcha_input.clear()
                    captcha_input.send_keys(captcha_text)

                    verify_button = wait_for_first(driver, VERIFY_BUTTON_LOCATORS, clickable=True, timeout=10)
                    if verify_button is None:
                        raise RuntimeError("VerifyButtonMissing")
                    stable_click(driver, verify_button)
                    time.sleep(1.8)

                    if check_server_down(driver):
                        raise RuntimeError("ServerDownAfterVerifyClick")

                    if is_visible_with_text(
                        driver,
                        UDYAM_FORMAT_ERROR_LOCATORS,
                        ["udyam should have first 5 letters", "invalid udyam", "wrong udyam"],
                    ):
                        wrong_udyam_detected = True
                        break

                    if is_visible_with_text(
                        driver,
                        UDYAM_REQUIRED_LOCATORS,
                        ["required"],
                    ):
                        try:
                            udyam_input = find_first(driver, UDYAM_INPUT_LOCATORS)
                            if udyam_input is not None:
                                udyam_input.clear()
                                udyam_input.send_keys(normalized_udyam)
                        except Exception:
                            pass
                        refresh_captcha_if_available(driver)
                        continue

                    if body_contains_any(
                        driver,
                        [
                            "wrong udyam number",
                            "invalid udyam",
                            "does not exist",
                        ],
                    ):
                        wrong_udyam_detected = True
                        break

                    status_message = read_status_message(driver).lower()
                    if status_message:
                        if any(
                            phrase in status_message
                            for phrase in [
                                "incorrect",
                                "invalid captcha",
                                "captcha code",
                                "captcha is incorrect",
                            ]
                        ) and "enterprise" not in status_message:
                            refresh_captcha_if_available(driver)
                            time.sleep(0.8)
                            continue

                    enterprise_name_element = wait_for_first(
                        driver,
                        ENTERPRISE_NAME_LOCATORS,
                        timeout=4,
                    )
                    if enterprise_name_element is not None and str(enterprise_name_element.text or "").strip():
                        captcha_ok = True
                        break

                    if body_contains_any(driver, ["incorrect captcha", "captcha is incorrect"]):
                        refresh_captcha_if_available(driver)
                        continue

                    refresh_captcha_if_available(driver)

                if wrong_udyam_detected:
                    print(f"[MSME single] Wrong Udyam Number: {normalized_udyam}")
                    return {}

                if not captcha_ok:
                    raise RuntimeError("CaptchaFailedMaxAttempts")

                enterprise_name_element = wait_for_first(
                    driver,
                    ENTERPRISE_NAME_LOCATORS,
                    timeout=12,
                )
                if enterprise_name_element is None:
                    raise RuntimeError("ResultPageNotLoadedAfterVerify")

                if CANCELLATION_MESSAGE.lower() in body_text(driver).lower():
                    print(f"[MSME single] Udyam is cancelled: {normalized_udyam}")
                    return {}

                major_activity = extract_major_activity(driver)
                enterprise_type = extract_classification_enterprise_type(driver)
                udyam_registration_date = extract_udyam_registration_date(driver)
                msme_category = normalize_msme_category(enterprise_type)
                msme_industry = normalize_msme_industry(major_activity)

                result = {
                    "MSME Category": msme_category,
                    "MSME Industry": msme_industry,
                    "Date of Udyam Registration": udyam_registration_date,
                    "Enterprise Type": enterprise_type,
                    "Major Activity": major_activity,
                }
                break
            except Exception as exc:
                if isinstance(exc, WebDriverException) and is_connection_refused_error(exc):
                    print("[MSME single] Browser session is unavailable")
                    return {}
                print(f"[MSME single] Error in attempt: {exc}")
                time.sleep(2)

        if not result:
            print(f"[MSME single] Failed to process Udyam: {normalized_udyam}")
        return result
    finally:
        try:
            driver.quit()
        except Exception:
            pass
