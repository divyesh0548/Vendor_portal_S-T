from __future__ import annotations

import hashlib
import io
import mimetypes
import re
from datetime import date
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from PIL import Image, ImageChops, ImageEnhance, ImageFilter, ImageOps


MAX_UPLOAD_BYTES = 20 * 1024 * 1024
TEXT_PREVIEW_LIMIT = 1200
MAX_OCR_IMAGE_VARIANTS = 16
MAX_ROTATING_OCR_IMAGE_VARIANTS = 8

PAN_REGEX = re.compile(r"\b[A-Z]{5}[0-9]{4}[A-Z]\b")
GST_REGEX = re.compile(r"\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b")
CIN_REGEX = re.compile(r"\b[LU][0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6}\b")
PAN_CORE_REGEX = re.compile(r"[A-Z]{5}[0-9]{4}[A-Z]")
GST_CORE_REGEX = re.compile(r"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]")
CIN_CORE_REGEX = re.compile(r"[LU][0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6}")
DATE_CANDIDATE_REGEX = re.compile(
    r"\b(?:[0-3]?[0-9][\-/\.][01]?[0-9][\-/\.](?:[12][0-9]{3}|[0-9]{2})|"
    r"(?:[12][0-9]{3})[\-/\.][01]?[0-9][\-/\.][0-3]?[0-9])\b"
)
SPACED_DATE_CANDIDATE_REGEX = re.compile(r"\b([0-3]?[0-9])\s+([01]?[0-9])\s+([12][0-9]{3}|[0-9]{2})\b")
UDYAM_REGEX = re.compile(r"\bUDYAM[-\s]?[A-Z0-9]{2}[-\s]?[0-9]{2}[-\s]?[0-9]{7}\b", re.IGNORECASE)
ESI_17_DIGIT_REGEX = re.compile(r"\b[0-9]{17}\b")
ESI_LABEL_VALUE_REGEX = re.compile(r"\b(?:ESI|ESIC)[\s:/#-]*(?:NO|NUMBER|CODE)?[\s:/#-]*([0-9]{10,17})\b", re.IGNORECASE)
PF_CODE_REGEX = re.compile(
    r"\b[A-Z]{2}[/\-\s]?[A-Z]{3}[/\-\s]?[0-9]{4,7}(?:[/\-\s]?[0-9]{1,3})?(?:[/\-\s]?[0-9]{1,7})?\b"
)
PF_LABEL_VALUE_REGEX = re.compile(
    r"\b(?:PF|EPF|EPFO|PROVIDENT\s+FUND|ESTABLISHMENT)\s*(?:NO|NUMBER|CODE|ID)?[\s:/#-]*([A-Z0-9][A-Z0-9/\-\s]{5,35})",
    re.IGNORECASE,
)
IFSC_REGEX = re.compile(r"\b[A-Z]{4}0[A-Z0-9]{6}\b")
IFSC_CORE_REGEX = re.compile(r"[A-Z]{4}[0O][A-Z0-9]{6}")
SWIFT_REGEX = re.compile(r"\b[A-Z]{4}[A-Z]{2}[A-Z0-9]{2}(?:[A-Z0-9]{3})?\b")
IBAN_CORE_REGEX = re.compile(r"[A-Z]{2}[0-9]{2}[A-Z0-9]{11,30}")
IBAN_LENGTH_BY_COUNTRY = {
    "AD": 24,
    "AE": 23,
    "AL": 28,
    "AT": 20,
    "AZ": 28,
    "BA": 20,
    "BE": 16,
    "BG": 22,
    "BH": 22,
    "BR": 29,
    "CH": 21,
    "CR": 22,
    "CY": 28,
    "CZ": 24,
    "DE": 22,
    "DK": 18,
    "DO": 28,
    "EE": 20,
    "ES": 24,
    "FI": 18,
    "FO": 18,
    "FR": 27,
    "GB": 22,
    "GE": 22,
    "GI": 23,
    "GL": 18,
    "GR": 27,
    "GT": 28,
    "HR": 21,
    "HU": 28,
    "IE": 22,
    "IL": 23,
    "IS": 26,
    "IT": 27,
    "JO": 30,
    "KW": 30,
    "KZ": 20,
    "LB": 28,
    "LC": 32,
    "LI": 21,
    "LT": 20,
    "LU": 20,
    "LV": 21,
    "MC": 27,
    "MD": 24,
    "ME": 22,
    "MK": 19,
    "MR": 27,
    "MT": 31,
    "MU": 30,
    "NL": 18,
    "NO": 15,
    "PK": 24,
    "PL": 28,
    "PS": 29,
    "PT": 25,
    "QA": 29,
    "RO": 24,
    "RS": 22,
    "SA": 24,
    "SC": 31,
    "SE": 24,
    "SI": 19,
    "SK": 24,
    "SM": 27,
    "TN": 24,
    "TR": 26,
    "UA": 29,
    "VA": 22,
    "VG": 24,
    "XK": 20,
}
BANK_ACCOUNT_LABEL_VALUE_REGEX = re.compile(
    r"\b(?:(?:SB|CA|CURRENT|SAVINGS)\s*)?(?:A\s*/?\s*C|AC|ACCT|ACCOUNT|BANK\s+ACCOUNT|BENEFICIARY\s+ACCOUNT)"
    r"\s*(?:NO|NUMBER|NUM|#)?\.?[\s:/#\-.]*([0-9][0-9\s-]{7,25}[0-9])\b",
    re.IGNORECASE,
)
BANK_ACCOUNT_ALPHANUMERIC_LABEL_VALUE_REGEX = re.compile(
    r"\b(?:(?:BENEFICIARY|BANK|FOREIGN|WIRE)\s*)?(?:A\s*/?\s*C|AC|ACCT|ACCOUNT)"
    r"\s*(?:NO|NUMBER|NUM|#)?\.?[\s:/#\-.]*([A-Z0-9][A-Z0-9\s\-/]{4,40})",
    re.IGNORECASE,
)
BANK_ACCOUNT_LABEL_REGEX = re.compile(
    r"\b(?:(?:SB|CA|CURRENT|SAVINGS)\s*)?(?:A\s*/?\s*C|ACCT|ACCOUNT|BANK\s+ACCOUNT|BENEFICIARY\s+ACCOUNT)"
    r"\s*(?:NO|NUMBER|NUM|#)?\.?\b",
    re.IGNORECASE,
)
BANK_ACCOUNT_NUMBER_CANDIDATE_REGEX = re.compile(r"\b[0-9][0-9\s-]{7,25}[0-9]\b")
BANK_ACCOUNT_ALPHANUMERIC_CANDIDATE_REGEX = re.compile(r"\b[A-Z0-9][A-Z0-9\-]{4,34}\b", re.IGNORECASE)

PAN_ENTITY_BY_FOURTH_CHAR = {
    "P": "Individual",
    "C": "Company",
    "H": "HUF",
    "A": "AOP",
    "B": "BOI",
    "G": "Government",
    "J": "Juridical Person",
    "L": "Local Authority",
    "F": "Firm",
    "T": "Trust",
}

MSME_CATEGORY_LABELS = {
    "micro": "Micro",
    "small": "Small",
    "medium": "Medium",
    "not an msme": "Not an MSME",
}

MSME_INDUSTRY_LABELS = {
    "manufacturing": "Manufacturing",
    "manufacturer": "Manufacturing",
    "service": "Service",
    "services": "Service",
    "trading": "Trading",
    "trade": "Trading",
}

KNOWN_BANK_NAME_NORMALIZATIONS = (
    ("unicredit", "UniCredit Bank"),
    ("icici", "ICICI Bank"),
    ("hdfc", "HDFC Bank"),
    ("axis", "Axis Bank"),
    ("state bank of india", "State Bank of India"),
    ("sbi", "State Bank of India"),
    ("kotak", "Kotak Mahindra Bank"),
    ("indusind", "IndusInd Bank"),
    ("idfc", "IDFC First Bank"),
    ("idbi", "IDBI Bank"),
    ("punjab national", "Punjab National Bank"),
    ("pnb", "Punjab National Bank"),
    ("canara", "Canara Bank"),
    ("union bank", "Union Bank"),
    ("central bank", "Central Bank of India"),
    ("indian bank", "Indian Bank"),
    ("federal bank", "Federal Bank"),
    ("rbl", "RBL Bank"),
    ("bandhan", "Bandhan Bank"),
    ("yes bank", "Yes Bank"),
    ("standard chartered", "Standard Chartered Bank"),
    ("hsbc", "HSBC"),
    ("dbs", "DBS Bank"),
    ("citi bank", "Citibank"),
    ("citibank", "Citibank"),
)

DOCUMENT_TYPE_ALIASES = {
    "gst": "gst",
    "gstcertificate": "gst",
    "gstcertificate_file": "gst",
    "gstcertificatefilename": "gst",
    "gstdocument": "gst",
    "gstattachment": "gst",
    "gstcertificatename": "gst",
    "pan": "pan",
    "panupload": "pan",
    "pandocument": "pan",
    "panattachment": "pan",
    "pancard": "pan",
    "pan_card_filename": "pan",
    "msme": "msme",
    "msmedocument": "msme",
    "msmecerti": "msme",
    "msme_certificate_filename": "msme",
    "udyam": "msme",
    "pf": "pf",
    "pfdocument": "pf",
    "pfcerti": "pf",
    "pf_certificate_filename": "pf",
    "cin": "cin",
    "cindocument": "cin",
    "cincerti": "cin",
    "incorporationcertificate": "cin",
    "incorporation_filename": "cin",
    "esi": "esi",
    "esidocument": "esi",
    "esicerti": "esi",
    "esi_filename": "esi",
    "bank": "bank",
    "bankdocument": "bank",
    "bank_document": "bank",
    "bankcertificate": "bank",
    "bankcerti": "bank",
    "cancelledcheque": "bank",
    "cancelledcheck": "bank",
    "cancelled_cheque_filename": "bank",
    "cancelledchequefilename": "bank",
    "bankattachment": "bank",
    "bankdocumentfile": "bank",
    "bankdocumentfilename": "bank",
    "bankdocumentattachment": "bank",
    "importbankdocument": "bank",
}


def register_document_extraction_routes(app: FastAPI) -> None:
    @app.post("/api/document-extraction/extract")
    async def extract_document_endpoint(
        file: UploadFile = File(...),
        documentType: str = Form(""),
        fieldName: str = Form(""),
    ) -> dict[str, Any]:
        filename = Path(str(file.filename or "document")).name
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="Uploaded document is empty")
        if len(content) > MAX_UPLOAD_BYTES:
            raise HTTPException(status_code=413, detail="Document is too large for extraction")

        expected_type = normalize_document_type(documentType or fieldName or filename)
        content_type = str(file.content_type or mimetypes.guess_type(filename)[0] or "").strip()
        return extract_document_details(
            content=content,
            filename=filename,
            content_type=content_type,
            expected_document_type=expected_type,
        )


def extract_document_details(
    *,
    content: bytes,
    filename: str = "",
    content_type: str = "",
    expected_document_type: str = "",
) -> dict[str, Any]:
    expected_type = normalize_document_type(expected_document_type or filename)
    text, extraction_method = extract_text_from_document(
        content,
        filename=filename,
        content_type=content_type,
        expected_document_type=expected_type,
    )
    canonical_fields = extract_fields_from_text(text, expected_document_type=expected_type)
    detected_types = detect_document_types(text, canonical_fields, filename)
    is_expected = determine_expected_document_match(expected_type, detected_types, canonical_fields)

    response_fields = build_form_field_aliases(canonical_fields)
    messages: list[str] = []
    if not text.strip():
        messages.append("No readable text found. Upload a clear PDF/image or enter the fields manually.")
    if expected_type and is_expected is False:
        detected_label = ", ".join(detected_types) if detected_types else "unknown"
        messages.append(f"Uploaded file does not look like a {expected_type.upper()} document. Detected: {detected_label}.")
    if not response_fields:
        messages.append("No supported document values were extracted.")

    return {
        "documentType": expected_type,
        "expectedDocumentType": expected_type,
        "detectedDocumentTypes": detected_types,
        "isExpectedDocument": is_expected,
        "fields": response_fields,
        "canonicalFields": canonical_fields,
        "confidence": score_confidence(expected_type, detected_types, canonical_fields, bool(text.strip())),
        "extractionMethod": extraction_method,
        "textPreview": collapse_spaces(text)[:TEXT_PREVIEW_LIMIT],
        "message": " ".join(messages),
    }


def normalize_document_type(value: Any) -> str:
    raw = str(value or "").strip().lower()
    if not raw:
        return ""
    key = re.sub(r"[^a-z0-9_]+", "", raw)
    if key in DOCUMENT_TYPE_ALIASES:
        return DOCUMENT_TYPE_ALIASES[key]

    for marker, normalized in DOCUMENT_TYPE_ALIASES.items():
        marker_key = re.sub(r"[^a-z0-9]+", "", marker)
        if marker_key and marker_key in key:
            return normalized
    return ""


def extract_text_from_document(
    content: bytes,
    *,
    filename: str = "",
    content_type: str = "",
    expected_document_type: str = "",
) -> tuple[str, str]:
    lowered_name = str(filename or "").lower()
    lowered_type = str(content_type or "").lower()
    is_pdf = lowered_name.endswith(".pdf") or "pdf" in lowered_type or content[:5] == b"%PDF-"
    is_image = lowered_type.startswith("image/") or lowered_name.endswith(
        (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff")
    )

    if is_pdf:
        candidates: list[tuple[str, str]] = []
        for extractor, method_name in (
            (_extract_pdf_text_with_pypdf, "pypdf"),
            (_extract_pdf_text_with_pymupdf, "pymupdf"),
            (_extract_pdf_text_with_pdfplumber, "pdfplumber"),
        ):
            text = extractor(content)
            if text and has_meaningful_extracted_text(text):
                candidates.append((text, method_name))
                if contains_supported_identifier(text):
                    return text, method_name

        text = _extract_pdf_text_with_pymupdf_ocr(content, expected_document_type=expected_document_type)
        if text and has_meaningful_extracted_text(text):
            candidates.append((text, "pymupdf-ocr"))
            if contains_supported_identifier(text):
                return combine_text_candidates(candidates), "pymupdf-ocr"

        text = _extract_pdf_text_rough(content)
        if text and has_meaningful_extracted_text(text):
            candidates.append((text, "pdf-rough"))

        best_text, best_method = choose_best_text_candidate(candidates)
        if best_text:
            return best_text, best_method
        return "", "pdf-unreadable"

    if is_image:
        text = _extract_image_text(content, expected_document_type=expected_document_type)
        return text, "image-ocr" if text else "image-unreadable"

    text = _decode_text_content(content)
    if text and has_meaningful_extracted_text(text):
        return text, "text"
    return "", "unsupported"


def _extract_pdf_text_with_pypdf(content: bytes) -> str:
    for module_name in ("pypdf", "PyPDF2"):
        try:
            module = __import__(module_name)
            reader = module.PdfReader(io.BytesIO(content))
            parts = []
            for page in getattr(reader, "pages", [])[:8]:
                page_text = page.extract_text() or ""
                if page_text:
                    parts.append(page_text)
            return "\n".join(parts)
        except Exception:
            continue
    return ""


def _extract_pdf_text_with_pymupdf(content: bytes) -> str:
    try:
        import fitz  # type: ignore

        parts: list[str] = []
        with fitz.open(stream=content, filetype="pdf") as document:
            for page_index in range(min(len(document), 8)):
                page = document.load_page(page_index)
                page_text = page.get_text("text") or ""
                if page_text:
                    parts.append(page_text)
        return "\n".join(parts)
    except Exception:
        return ""


def _extract_pdf_text_with_pdfplumber(content: bytes) -> str:
    try:
        import pdfplumber  # type: ignore

        parts: list[str] = []
        with pdfplumber.open(io.BytesIO(content)) as document:
            for page in document.pages[:8]:
                page_text = page.extract_text() or ""
                if page_text:
                    parts.append(page_text)
        return "\n".join(parts)
    except Exception:
        return ""


def _extract_pdf_text_with_pymupdf_ocr(content: bytes, *, expected_document_type: str = "") -> str:
    try:
        import fitz  # type: ignore

        candidates: list[tuple[str, str]] = []
        with fitz.open(stream=content, filetype="pdf") as document:
            for page_index in range(min(len(document), 3)):
                page = document.load_page(page_index)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
                image = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
                page_text = _ocr_pil_image(image, expected_document_type=expected_document_type, try_rotations=False)
                if page_text:
                    candidates.append((page_text, f"page-{page_index + 1}"))
        return combine_text_candidates(candidates)
    except Exception:
        return ""


def _extract_pdf_text_rough(content: bytes) -> str:
    raw = content.decode("latin-1", errors="ignore")
    if not raw:
        return ""

    candidates: list[str] = []
    for pattern in (
        r"\(([^()]{2,120})\)\s*Tj",
        r"\(([^()]{2,120})\)",
        r"<([0-9A-Fa-f]{8,240})>\s*Tj",
    ):
        for match in re.finditer(pattern, raw):
            token = match.group(1)
            if re.fullmatch(r"[0-9A-Fa-f]+", token or "") and len(token) % 2 == 0:
                try:
                    token = bytes.fromhex(token).decode("utf-16-be", errors="ignore")
                except Exception:
                    token = ""
            token = _unescape_pdf_string(token)
            if looks_like_text_token(token):
                candidates.append(token)
    return "\n".join(candidates)


def _unescape_pdf_string(value: str) -> str:
    text = str(value or "")
    text = text.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")
    text = re.sub(r"\\([nrtbf])", " ", text)
    text = re.sub(r"\\[0-7]{1,3}", " ", text)
    return text


def _decode_text_content(content: bytes) -> str:
    for encoding in ("utf-8", "utf-8-sig", "utf-16", "latin-1"):
        try:
            text = content.decode(encoding, errors="strict")
        except UnicodeDecodeError:
            continue
        if has_meaningful_extracted_text(text):
            return text
    return ""


def _extract_image_text(content: bytes, *, expected_document_type: str = "") -> str:
    try:
        image = ImageOps.exif_transpose(Image.open(io.BytesIO(content))).convert("RGB")
    except Exception:
        return ""
    return _ocr_pil_image(image, expected_document_type=expected_document_type)


def _ocr_pil_image(image: Image.Image, *, expected_document_type: str = "", try_rotations: bool = True) -> str:
    candidates: list[tuple[str, str]] = []
    angles = (0, 90, 270, 180) if try_rotations else (0,)
    for angle in angles:
        rotated = rotate_ocr_image(image, angle)
        variants = build_ocr_image_variants(rotated)
        if try_rotations:
            variants = variants[:MAX_ROTATING_OCR_IMAGE_VARIANTS]
        for variant_index, candidate in enumerate(variants):
            text = _ocr_with_tesseract(candidate)
            if not text or not has_meaningful_extracted_text(text):
                continue
            candidates.append((text, f"tesseract-{angle}-{variant_index}"))
            if len(candidates) >= 3:
                combined_so_far = combine_text_candidates(top_text_candidates(candidates))
                if is_strong_ocr_text(combined_so_far, expected_document_type=expected_document_type):
                    break

        combined_text = combine_text_candidates(candidates)
        if is_strong_ocr_text(combined_text, expected_document_type=expected_document_type):
            break
        if angle == 0 and try_rotations and not should_try_more_ocr_angles(combined_text, expected_document_type):
            break

    combined_text = combine_text_candidates(candidates)
    if (not candidates or not contains_supported_identifier(combined_text)) and should_try_easyocr_fallback(
        combined_text,
        expected_document_type,
    ):
        for angle in angles:
            easy_text = _ocr_with_easyocr(rotate_ocr_image(image, angle))
            if easy_text and has_meaningful_extracted_text(easy_text):
                candidates.append((easy_text, f"easyocr-{angle}"))
                if is_strong_ocr_text(easy_text, expected_document_type=expected_document_type):
                    break
    best_text, _ = choose_best_text_candidate(candidates)
    if not best_text:
        return ""
    return combine_text_candidates(top_text_candidates(candidates))


def should_try_more_ocr_angles(text: str, expected_document_type: str = "") -> bool:
    cleaned = collapse_spaces(text)
    if len(cleaned) < 80:
        return True
    return has_expected_document_hint(text, expected_document_type)


def should_try_easyocr_fallback(text: str, expected_document_type: str = "") -> bool:
    if not collapse_spaces(text):
        return False
    return has_expected_document_hint(text, expected_document_type)


def has_expected_document_hint(text: str, expected_document_type: str = "") -> bool:
    expected_type = normalize_document_type(expected_document_type)
    if expected_type == "pan":
        return looks_like_pan_document(text)
    if expected_type == "bank":
        return has_strong_bank_document_marker(text)
    if expected_type == "gst":
        return looks_like_gst_document(text)
    if expected_type == "msme":
        return looks_like_msme_document(text)
    if expected_type == "cin":
        return "certificate of incorporation" in str(text or "").lower() or bool(find_cin_number(re.sub(r"[^A-Z0-9]", "", str(text or "").upper())))
    if expected_type == "pf":
        return any(term in str(text or "").lower() for term in ("provident fund", "epfo", "employee provident"))
    if expected_type == "esi":
        return any(term in str(text or "").lower() for term in ("employees' state insurance", "employee state insurance", "esic"))
    return contains_supported_identifier(text)


def top_text_candidates(candidates: list[tuple[str, str]], limit: int = 4) -> list[tuple[str, str]]:
    return sorted(candidates, key=lambda item: score_text_candidate(item[0]), reverse=True)[:limit]


def rotate_ocr_image(image: Image.Image, angle: int) -> Image.Image:
    if not angle:
        return image.convert("RGB")
    return image.convert("RGB").rotate(angle, expand=True, fillcolor="white")


def build_ocr_image_variants(image: Image.Image) -> list[Image.Image]:
    base = resize_ocr_image(image.convert("RGB"))
    sources = [base]
    trimmed = trim_ocr_image_borders(base)
    if trimmed is not None:
        sources.append(trimmed)

    variants: list[Image.Image] = []
    for source in sources:
        gray = ImageOps.autocontrast(source.convert("L"))
        contrast = ImageEnhance.Contrast(gray).enhance(2.2)
        brightened = ImageEnhance.Brightness(contrast).enhance(1.25)
        darkened = ImageEnhance.Brightness(contrast).enhance(0.75)
        denoised = gray.filter(ImageFilter.MedianFilter(size=3))
        sharpened = contrast.filter(ImageFilter.UnsharpMask(radius=2, percent=180, threshold=3))
        threshold = contrast.point(lambda px: 0 if px < 170 else 255, mode="L")
        light_threshold = brightened.point(lambda px: 0 if px < 185 else 255, mode="L")
        dark_threshold = darkened.point(lambda px: 0 if px < 145 else 255, mode="L")
        inverted = ImageOps.invert(gray)
        deskewed = deskew_ocr_image(gray)
        cv2_variants = build_cv2_ocr_image_variants(gray)

        variants.extend(
            [
                source,
                gray,
                contrast,
                sharpened,
                threshold,
                light_threshold,
                *cv2_variants,
                denoised,
                dark_threshold,
                inverted,
            ]
        )
        if deskewed is not None:
            variants.extend([deskewed, *build_cv2_ocr_image_variants(deskewed)])
    return dedupe_ocr_image_variants(variants)[:MAX_OCR_IMAGE_VARIANTS]


def resize_ocr_image(image: Image.Image) -> Image.Image:
    width, height = image.size
    long_edge = max(width, height)
    if long_edge <= 0:
        return image.convert("RGB")
    target_long_edge = 2600 if long_edge < 1100 else 2200 if long_edge < 1800 else long_edge
    scale = target_long_edge / long_edge
    if scale <= 1.05:
        return image.convert("RGB")
    scale = min(scale, 4.0)
    resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
    return image.resize((max(1, int(width * scale)), max(1, int(height * scale))), resampling).convert("RGB")


def trim_ocr_image_borders(image: Image.Image) -> Image.Image | None:
    try:
        gray = ImageOps.autocontrast(image.convert("L"))
        white = Image.new("L", gray.size, 255)
        diff = ImageChops.difference(gray, white)
        mask = diff.point(lambda px: 255 if px > 18 else 0)
        bbox = mask.getbbox()
        if not bbox:
            return None
        left, top, right, bottom = bbox
        width, height = image.size
        crop_width = right - left
        crop_height = bottom - top
        if crop_width < 80 or crop_height < 80:
            return None
        if crop_width > width * 0.96 and crop_height > height * 0.96:
            return None
        pad_x = max(12, int(crop_width * 0.04))
        pad_y = max(12, int(crop_height * 0.04))
        box = (
            max(0, left - pad_x),
            max(0, top - pad_y),
            min(width, right + pad_x),
            min(height, bottom + pad_y),
        )
        return image.crop(box).convert("RGB")
    except Exception:
        return None


def dedupe_ocr_image_variants(images: list[Image.Image]) -> list[Image.Image]:
    output: list[Image.Image] = []
    seen: set[tuple[str, tuple[int, int], str]] = set()
    for image in images:
        converted = image.convert("RGB")
        key = (converted.mode, converted.size, hashlib.blake2b(converted.tobytes(), digest_size=8).hexdigest())
        if key in seen:
            continue
        seen.add(key)
        output.append(converted)
    return output


def build_cv2_ocr_image_variants(image: Image.Image) -> list[Image.Image]:
    try:
        import cv2  # type: ignore
        import numpy as np

        gray = np.array(image.convert("L"))
        denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(denoised)
        adaptive = cv2.adaptiveThreshold(
            clahe,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31,
            11,
        )
        _, otsu = cv2.threshold(clahe, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return [Image.fromarray(clahe), Image.fromarray(adaptive), Image.fromarray(otsu)]
    except Exception:
        return []


def deskew_ocr_image(image: Image.Image) -> Image.Image | None:
    try:
        import cv2  # type: ignore
        import numpy as np

        gray = np.array(image.convert("L"))
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        coords = np.column_stack(np.where(binary > 0))
        if len(coords) < 80:
            return None
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle
        if abs(angle) < 0.5 or abs(angle) > 18:
            return None
        height, width = gray.shape[:2]
        matrix = cv2.getRotationMatrix2D((width // 2, height // 2), angle, 1.0)
        rotated = cv2.warpAffine(gray, matrix, (width, height), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
        return Image.fromarray(rotated)
    except Exception:
        return None


def _ocr_with_tesseract(image: Image.Image) -> str:
    try:
        import pytesseract

        configure_tesseract_if_available(pytesseract)
        text_parts: list[str] = []
        seen_lines: set[str] = set()
        for config in ("--oem 3 --psm 6", "--oem 3 --psm 11"):
            text = pytesseract.image_to_string(image, config=config) or ""
            for line in text.splitlines():
                cleaned = collapse_spaces(line)
                if not cleaned:
                    continue
                dedupe_key = cleaned.lower()
                if dedupe_key in seen_lines:
                    continue
                seen_lines.add(dedupe_key)
                text_parts.append(cleaned)
        return "\n".join(text_parts)
    except Exception:
        return ""


def configure_tesseract_if_available(pytesseract_module: Any) -> None:
    current_cmd = str(getattr(pytesseract_module.pytesseract, "tesseract_cmd", "") or "")
    if current_cmd and current_cmd.lower() != "tesseract":
        return

    for candidate in (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\Maulik Khunt\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
    ):
        if Path(candidate).exists():
            pytesseract_module.pytesseract.tesseract_cmd = candidate
            return


_easyocr_reader: Any | None = None


def _ocr_with_easyocr(image: Image.Image) -> str:
    global _easyocr_reader
    try:
        import numpy as np
        import easyocr

        if _easyocr_reader is None:
            _easyocr_reader = easyocr.Reader(["en"], gpu=False)
        results = _easyocr_reader.readtext(np.array(image.convert("RGB")), detail=0, paragraph=True)
        return "\n".join(str(item or "").strip() for item in results if str(item or "").strip())
    except Exception:
        return ""


def has_meaningful_extracted_text(text: str) -> bool:
    cleaned = collapse_spaces(text)
    if len(cleaned) < 12:
        return False
    if len(cleaned) > 40:
        readable_count = sum(1 for ch in cleaned if ch.isascii() and (ch.isalnum() or ch.isspace() or ch in "-/:.,()&'"))
        if readable_count / max(len(cleaned), 1) < 0.55:
            return False
    return any(ch.isalpha() for ch in cleaned) or any(ch.isdigit() for ch in cleaned)


def combine_text_candidates(candidates: list[tuple[str, str]]) -> str:
    lines: list[str] = []
    seen: set[str] = set()
    for text, _method in candidates:
        for line in str(text or "").replace("\r", "\n").split("\n"):
            cleaned = collapse_spaces(line)
            if not cleaned:
                continue
            key = re.sub(r"[^a-z0-9]+", "", cleaned.lower())
            if key and key in seen:
                continue
            if key:
                seen.add(key)
            lines.append(cleaned)
    return "\n".join(lines)


def choose_best_text_candidate(candidates: list[tuple[str, str]]) -> tuple[str, str]:
    if not candidates:
        return "", ""
    scored = [
        (score_text_candidate(text), index, text, method)
        for index, (text, method) in enumerate(candidates)
        if has_meaningful_extracted_text(text)
    ]
    if not scored:
        return "", ""
    scored.sort(key=lambda item: (item[0], -item[1]), reverse=True)
    score, _index, text, method = scored[0]
    if score <= 0:
        return "", ""
    return text, method


def score_text_candidate(text: str) -> float:
    if not has_meaningful_extracted_text(text):
        return 0.0
    score = min(10.0, len(collapse_spaces(text)) / 120)
    try:
        fields = extract_fields_from_text(text)
        score += len([value for value in fields.values() if value]) * 18
    except Exception:
        fields = {}
    try:
        if contains_supported_identifier(text):
            score += 35
    except Exception:
        pass
    lowered = str(text or "").lower()
    if any(
        term in lowered
        for term in ("gstin", "permanent account", "udyam", "ifsc", "bank account", "bank certificate", "cancelled cheque")
    ):
        score += 8
    if fields.get("bankAccountNumber") and fields.get("bankName"):
        score += 15
    return score


def is_strong_ocr_text(text: str, *, expected_document_type: str = "") -> bool:
    if not text or not contains_supported_identifier(text):
        return False
    try:
        fields = extract_fields_from_text(text, expected_document_type=expected_document_type)
    except Exception:
        return True
    return len([value for value in fields.values() if value]) >= 2


def looks_like_text_token(value: str) -> bool:
    text = collapse_spaces(value)
    if len(text) < 2:
        return False
    printable_count = sum(1 for ch in text if ch.isprintable())
    if printable_count < max(2, int(len(text) * 0.8)):
        return False
    return any(ch.isalnum() for ch in text)


def collapse_spaces(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\x00", " ")).strip()


def normalize_for_identifier_scan(text: str) -> str:
    normalized = str(text or "").upper()
    normalized = normalized.replace("O", "0")
    return normalized


def contains_supported_identifier(text: str) -> bool:
    compact_upper = re.sub(r"[^A-Z0-9]", "", str(text or "").upper())
    if find_gst_number(compact_upper):
        return True
    if find_pan_number(compact_upper, text):
        return True
    if find_cin_number(compact_upper):
        return True
    if find_udyam_number(text):
        return True
    if find_esi_number(text):
        return True
    if find_pf_number(text):
        return True
    if find_ifsc_code(text):
        return True
    if find_swift_code(text):
        return True
    if find_iban(text):
        return True
    if find_bank_account_number(text):
        return True
    return False


def extract_fields_from_text(text: str, *, expected_document_type: str = "") -> dict[str, str]:
    lines = get_clean_lines(text)
    joined = "\n".join(lines)
    compact_upper = re.sub(r"[^A-Z0-9]", "", str(text or "").upper())
    expected_type = normalize_document_type(expected_document_type)

    fields: dict[str, str] = {}

    gst_number = find_gst_number(compact_upper)
    if gst_number:
        fields["gstNumber"] = gst_number
        pan_from_gst = gst_number[2:12]
        if PAN_REGEX.fullmatch(pan_from_gst):
            fields["panNumber"] = pan_from_gst

    pan_number = find_pan_number(compact_upper, text)
    if pan_number and not (gst_number and fields.get("panNumber") and pan_number != fields["panNumber"]):
        fields["panNumber"] = pan_number

    cin_number = find_cin_number(compact_upper)
    if cin_number:
        fields["cinNumber"] = cin_number

    udyam_number = find_udyam_number(text)
    if udyam_number:
        fields["udyamNumber"] = udyam_number

    if expected_type in ("", "pf") or (expected_type != "bank" and "provident" in joined.lower()):
        pf_number = find_pf_number(joined)
        if pf_number:
            fields["pfNumber"] = pf_number

    if expected_type in ("", "esi") or (expected_type != "bank" and "esic" in joined.lower()):
        esi_number = find_esi_number(joined)
        if esi_number:
            fields["esiNumber"] = esi_number

    ifsc_code = find_ifsc_code(joined)
    if ifsc_code:
        fields["ifscCode"] = ifsc_code

    swift_code = find_swift_code(joined)
    if swift_code:
        fields["swiftCode"] = swift_code

    iban = find_iban(joined)
    if iban:
        fields["iban"] = iban

    should_extract_bank = should_extract_bank_fields(joined, fields, expected_type)
    if should_extract_bank:
        bank_account_number = find_bank_account_number(joined)
        if bank_account_number:
            fields["bankAccountNumber"] = bank_account_number

        beneficiary_name = find_beneficiary_name(lines)
        if beneficiary_name:
            fields["beneficiaryName"] = beneficiary_name

        bank_name = find_bank_name(lines)
        if bank_name:
            fields["bankName"] = bank_name

        branch_name = find_branch_name(lines)
        if branch_name:
            fields["branchName"] = branch_name

        bank_document_date = find_bank_document_date(lines)
        if bank_document_date:
            fields["bankDocumentDate"] = bank_document_date
            fields["chequeDate"] = bank_document_date
    else:
        fields.pop("ifscCode", None)
        fields.pop("swiftCode", None)
        fields.pop("iban", None)

    if looks_like_gst_document(joined) or fields.get("gstNumber"):
        trade_name_gst = find_label_value(
            lines,
            ("trade name", "trade name, if any", "legal name", "legal name of business", "name of business"),
            forbidden=("father", "date", "address", "status", "type", "constitution"),
        )
        if trade_name_gst:
            fields["tradeNameGst"] = trade_name_gst

        taxpayer_type = find_label_value(
            lines,
            ("taxpayer type", "constitution of business", "type of taxpayer"),
            forbidden=("date", "address", "status"),
        )
        if taxpayer_type:
            fields["taxpayerTypeGst"] = normalize_title_value(taxpayer_type)

    if looks_like_pan_document(joined) or (fields.get("panNumber") and not fields.get("gstNumber")):
        trade_name_pan = find_pan_name(lines, fields.get("panNumber", ""))
        if trade_name_pan:
            fields["tradeNamePan"] = trade_name_pan

        pan_date = find_pan_date(lines)
        if pan_date:
            fields["panDate"] = pan_date

    if fields.get("panNumber"):
        entity_type = derive_business_entity_from_pan(fields["panNumber"])
        if entity_type:
            fields["businessEntityType"] = entity_type
        if fields.get("panDate"):
            if entity_type == "Individual":
                fields["dateOfBirth"] = fields["panDate"]
            else:
                fields["dateOfIncorporation"] = fields["panDate"]

    if looks_like_msme_document(joined) or fields.get("udyamNumber"):
        msme_category = find_msme_category(joined)
        if msme_category:
            fields["msmeCategory"] = msme_category

        msme_industry = find_msme_industry(joined)
        if msme_industry:
            fields["msmeIndustry"] = msme_industry

    return fields


def should_extract_bank_fields(text: str, fields: dict[str, str], expected_type: str = "") -> bool:
    normalized_expected = normalize_document_type(expected_type)
    has_bank_marker = has_strong_bank_document_marker(text)
    has_bank_identifier = any(fields.get(key) for key in ("ifscCode", "swiftCode", "iban"))
    if normalized_expected == "bank":
        return has_bank_marker or has_bank_identifier or has_bank_account_label_text(text)
    if normalized_expected and normalized_expected != "bank" and not has_bank_marker:
        return False
    return has_bank_marker or has_bank_identifier


def get_clean_lines(text: str) -> list[str]:
    lines: list[str] = []
    for raw_line in str(text or "").replace("\r", "\n").split("\n"):
        line = collapse_spaces(raw_line)
        if not line:
            continue
        lines.append(line)
    return lines


def find_pan_date(lines: list[str]) -> str:
    date_labels = (
        "date of birth",
        "dob",
        "birth",
        "date of incorporation",
        "incorporation",
    )
    candidates: list[tuple[int, int, str]] = []
    for index, line in enumerate(lines):
        nearby = " ".join(lines[max(0, index - 2) : index + 3])
        nearby_lower = nearby.lower()
        line_lower = line.lower()
        has_label = any(label in line_lower or label in nearby_lower for label in date_labels)
        date_value = find_first_date_value(line, allow_compact=has_label)
        if not date_value and has_label:
            date_value = find_first_date_value(nearby, allow_compact=True)
        if not date_value:
            continue
        score = 10
        if has_label:
            score += 40
        if any(token in nearby_lower for token in ("permanent account", "income tax", "pan")):
            score += 10
        if any(token in nearby_lower for token in ("signature", "father", "mother")):
            score -= 5
        candidates.append((score, -index, date_value))
    if not candidates:
        return ""
    candidates.sort(reverse=True)
    return candidates[0][2]


def find_bank_document_date(lines: list[str]) -> str:
    candidates: list[tuple[int, int, str]] = []
    for index, line in enumerate(lines):
        nearby = " ".join(lines[max(0, index - 2) : index + 3])
        nearby_lower = nearby.lower()
        if not any(token in nearby_lower for token in ("date", "cheque", "check", "issued", "booking")):
            continue
        date_value = find_first_date_value(line) or find_first_date_value(nearby)
        if not date_value:
            continue
        score = 10
        if "cheque" in nearby_lower or "check" in nearby_lower:
            score += 20
        if "date" in line.lower():
            score += 15
        if any(token in nearby_lower for token in ("booking", "statement", "period")):
            score -= 5
        candidates.append((score, -index, date_value))
    if not candidates:
        return ""
    candidates.sort(reverse=True)
    return candidates[0][2]


def find_first_date_value(text: str, *, allow_compact: bool = False) -> str:
    for value in find_date_values(text, allow_compact=allow_compact):
        return value
    return ""


def find_date_values(text: str, *, allow_compact: bool = False) -> list[str]:
    output: list[str] = []
    seen: set[str] = set()

    def add(raw_value: str) -> None:
        normalized = normalize_date_value(raw_value)
        if normalized and normalized not in seen:
            seen.add(normalized)
            output.append(normalized)

    raw_text = str(text or "")
    for match in DATE_CANDIDATE_REGEX.finditer(raw_text):
        add(match.group(0))
    for match in SPACED_DATE_CANDIDATE_REGEX.finditer(raw_text):
        add("/".join(match.groups()))
    if allow_compact:
        for match in re.finditer(r"\b[0-9]{8,10}\b", raw_text):
            add(match.group(0))
    return output


def normalize_date_value(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    parts = [part for part in re.split(r"[\s\-/\.]+", text) if part]
    if len(parts) == 3:
        if len(parts[0]) == 4:
            year_value, month_value, day_value = parts
        else:
            day_value, month_value, year_value = parts
        return build_iso_date(day_value, month_value, year_value)

    digits = re.sub(r"\D", "", text)
    if len(digits) == 8:
        for day_value, month_value, year_value in (
            (digits[:2], digits[2:4], digits[4:]),
            (digits[6:], digits[4:6], digits[:4]),
        ):
            normalized = build_iso_date(day_value, month_value, year_value)
            if normalized:
                return normalized
    if len(digits) in (9, 10):
        for index in range(2, len(digits) - 4):
            shortened = digits[:index] + digits[index + 1 :]
            normalized = normalize_date_value(shortened)
            if normalized:
                return normalized
    return ""


def build_iso_date(day_value: str, month_value: str, year_value: str) -> str:
    try:
        day_int = int(day_value)
        month_int = int(month_value)
        year_int = int(year_value)
        if year_int < 100:
            year_int += 2000 if year_int <= 49 else 1900
        if year_int < 1900 or year_int > 2100:
            return ""
        return date(year_int, month_int, day_int).isoformat()
    except Exception:
        return ""


def find_gst_number(compact_upper_text: str) -> str:
    compact = str(compact_upper_text or "")
    candidates: list[tuple[float, int, str]] = []
    for match in GST_CORE_REGEX.finditer(compact_upper_text):
        candidates.append((80 + identifier_context_score(compact, match.start(), gst_context_keywords()), match.start(), match.group(0)))
    for index in range(0, max(0, len(compact) - 14)):
        candidate = normalize_gst_candidate(compact[index : index + 15])
        if GST_CORE_REGEX.fullmatch(candidate):
            raw = compact[index : index + 15]
            correction_penalty = sum(1 for raw_ch, normalized_ch in zip(raw, candidate) if raw_ch != normalized_ch) * 2
            score = identifier_context_score(compact, index, gst_context_keywords()) - correction_penalty
            if score > 0:
                candidates.append((score, index, candidate))
    if not candidates:
        return ""
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][2]


def find_pan_number(compact_upper_text: str, original_text: str = "") -> str:
    compact = str(compact_upper_text or "")
    has_document_context = looks_like_pan_document(original_text)
    ignored_spans = (
        get_ifsc_spans_in_compact(compact)
        + get_iban_spans_in_compact(compact)
        + get_labeled_swift_spans_in_compact(compact, original_text)
        + get_labeled_bank_account_spans_in_compact(compact, original_text)
    )
    candidates: list[tuple[float, int, str]] = []
    for match in PAN_CORE_REGEX.finditer(compact_upper_text):
        if spans_overlap(match.start(), match.end(), ignored_spans):
            continue
        pan = match.group(0)
        if pan[3] in PAN_ENTITY_BY_FOURTH_CHAR:
            context_score = identifier_context_score(compact, match.start(), pan_context_keywords())
            if context_score > 0 or has_document_context:
                candidates.append((20 + context_score, match.start(), pan))
    for index in range(0, max(0, len(compact) - 9)):
        if spans_overlap(index, index + 10, ignored_spans):
            continue
        pan = normalize_pan_candidate(compact[index : index + 10])
        if PAN_CORE_REGEX.fullmatch(pan) and pan[3] in PAN_ENTITY_BY_FOURTH_CHAR:
            raw = compact[index : index + 10]
            correction_penalty = sum(1 for raw_ch, pan_ch in zip(raw, pan) if raw_ch != pan_ch) * 2
            score = identifier_context_score(compact, index, pan_context_keywords()) - correction_penalty
            if score > 0 or (has_document_context and correction_penalty <= 2):
                candidates.append((score, index, pan))
    if not candidates:
        return ""
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][2]
    return ""


def get_ifsc_spans_in_compact(compact_upper_text: str) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    compact = str(compact_upper_text or "").upper()
    for match in re.finditer(rf"(?=({IFSC_CORE_REGEX.pattern}))", compact):
        candidate = normalize_ifsc_candidate(match.group(1))
        if is_plausible_ifsc_code(candidate):
            spans.append((match.start(1), match.end(1)))
    return spans


def get_iban_spans_in_compact(compact_upper_text: str) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    compact = str(compact_upper_text or "").upper()
    for match in re.finditer(rf"(?=({IBAN_CORE_REGEX.pattern}))", compact):
        start = match.start(1)
        raw = normalize_iban_candidate(match.group(1))
        for length in range(min(34, len(raw)), 14, -1):
            candidate = raw[:length]
            if is_valid_iban(candidate):
                spans.append((start, start + length))
                break
    return spans


def get_labeled_bank_account_spans_in_compact(compact_upper_text: str, original_text: str = "") -> list[tuple[int, int]]:
    compact = str(compact_upper_text or "").upper()
    spans: list[tuple[int, int]] = []
    if not compact or not original_text:
        return spans

    for line in get_clean_lines(original_text):
        if not has_bank_account_label_text(line):
            continue
        for pattern in (BANK_ACCOUNT_NUMBER_CANDIDATE_REGEX, BANK_ACCOUNT_ALPHANUMERIC_CANDIDATE_REGEX):
            for match in pattern.finditer(line):
                candidate = normalize_bank_account_number(match.group(0))
                if not is_plausible_bank_account_number(candidate):
                    continue
                for found in re.finditer(re.escape(candidate), compact):
                    spans.append((found.start(), found.end()))
    return spans


def get_labeled_swift_spans_in_compact(compact_upper_text: str, original_text: str = "") -> list[tuple[int, int]]:
    compact = str(compact_upper_text or "").upper()
    spans: list[tuple[int, int]] = []
    if not compact or not original_text:
        return spans

    for line in get_clean_lines(original_text):
        if not has_swift_label_text(line):
            continue
        for candidate in extract_swift_candidates_from_text(line):
            normalized = normalize_swift_candidate(candidate)
            if not is_plausible_swift_code(normalized):
                continue
            for found in re.finditer(re.escape(normalized), compact):
                spans.append((found.start(), found.end()))
    return spans


def spans_overlap(start: int, end: int, spans: list[tuple[int, int]]) -> bool:
    return any(start < span_end and end > span_start for span_start, span_end in spans)


def pan_context_keywords() -> tuple[str, ...]:
    return (
        "PERMANENTACCOUNTNUMBER",
        "ACCQUALNUMBER",
        "PANNUMBER",
        "PANNO",
    )


def gst_context_keywords() -> tuple[str, ...]:
    return (
        "GSTIN",
        "GSTNUMBER",
        "REGISTRATIONNUMBER",
        "REGISTRATIONNO",
        "GSTREGISTRATION",
        "GOODSANDSERVICESTAX",
        "GST",
    )


def identifier_context_score(compact_text: str, index: int, keywords: tuple[str, ...]) -> float:
    best_score = 0.0
    compact = str(compact_text or "")
    for keyword in keywords:
        position = compact.rfind(keyword, 0, max(index, 0))
        if position < 0:
            continue
        distance = index - (position + len(keyword))
        if 0 <= distance <= 80:
            best_score = max(best_score, 60 - distance)
    return best_score


def find_cin_number(compact_upper_text: str) -> str:
    match = CIN_CORE_REGEX.search(compact_upper_text)
    return match.group(0) if match else ""


def normalize_pan_candidate(value: str) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    if len(raw) != 10:
        return raw
    chars = list(raw)
    for index in (0, 1, 2, 3, 4, 9):
        chars[index] = ocr_char_to_letter(chars[index])
    for index in (5, 6, 7, 8):
        chars[index] = ocr_char_to_digit(chars[index])
    return "".join(chars)


def normalize_gst_candidate(value: str) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    if len(raw) != 15:
        return raw
    chars = list(raw)
    for index in (0, 1, 7, 8, 9, 10):
        chars[index] = ocr_char_to_digit(chars[index])
    for index in (2, 3, 4, 5, 6, 11):
        chars[index] = ocr_char_to_letter(chars[index])
    chars[13] = "Z" if chars[13] in {"2", "7"} else ocr_char_to_letter(chars[13])
    return "".join(chars)


def ocr_char_to_digit(ch: str) -> str:
    return {
        "O": "0",
        "Q": "0",
        "D": "0",
        "I": "1",
        "L": "1",
        "T": "1",
        "Z": "2",
        "S": "5",
        "B": "8",
        "G": "6",
    }.get(str(ch or "").upper(), str(ch or "").upper())


def ocr_char_to_letter(ch: str) -> str:
    return {
        "0": "O",
        "1": "I",
        "2": "Z",
        "5": "S",
        "6": "G",
        "8": "B",
    }.get(str(ch or "").upper(), str(ch or "").upper())


def find_udyam_number(text: str) -> str:
    match = UDYAM_REGEX.search(str(text or ""))
    if not match:
        return ""
    raw = re.sub(r"[\s-]+", "-", match.group(0).strip().upper())
    parts = raw.split("-")
    if len(parts) == 4:
        return f"UDYAM-{parts[1]}-{parts[2]}-{parts[3]}"
    compact = re.sub(r"[^A-Z0-9]", "", raw)
    if compact.startswith("UDYAM") and len(compact) >= 16:
        return f"UDYAM-{compact[5:7]}-{compact[7:9]}-{compact[9:16]}"
    return raw


def find_pf_number(text: str) -> str:
    labeled = PF_LABEL_VALUE_REGEX.search(str(text or ""))
    if labeled:
        candidate = clean_identifier_value(labeled.group(1), allowed="/-")
        candidate = trim_identifier(candidate)
        if is_plausible_pf_number(candidate):
            return candidate

    for match in PF_CODE_REGEX.finditer(str(text or "").upper()):
        candidate = clean_identifier_value(match.group(0), allowed="/-")
        if is_plausible_pf_number(candidate):
            return candidate
    return ""


def is_plausible_pf_number(value: str) -> bool:
    candidate = str(value or "").strip().upper()
    if len(candidate) < 6 or len(candidate) > 35:
        return False
    if PAN_REGEX.fullmatch(re.sub(r"[^A-Z0-9]", "", candidate)):
        return False
    if GST_REGEX.fullmatch(re.sub(r"[^A-Z0-9]", "", candidate)):
        return False
    return any(ch.isdigit() for ch in candidate) and any(ch.isalpha() for ch in candidate)


def find_esi_number(text: str) -> str:
    labeled = ESI_LABEL_VALUE_REGEX.search(str(text or ""))
    if labeled:
        return clean_identifier_value(labeled.group(1))

    for match in ESI_17_DIGIT_REGEX.finditer(str(text or "")):
        return match.group(0)
    return ""


def find_ifsc_code(text: str) -> str:
    raw_upper = str(text or "").upper()
    lines = get_clean_lines(raw_upper)
    for index, line in enumerate(lines):
        if "IFSC" not in line and "IFS CODE" not in line and "IFSCCODE" not in re.sub(r"[^A-Z0-9]", "", line):
            continue
        candidate_text = line
        if index + 1 < len(lines):
            candidate_text = f"{candidate_text} {lines[index + 1]}"
        for candidate in extract_ifsc_candidates_from_text(candidate_text):
            if is_plausible_ifsc_code(candidate):
                return candidate

    candidates: list[tuple[float, int, str]] = []
    for line_index, line in enumerate(lines):
        if has_bank_account_label_text(line) and not any(token in line for token in ("IFSC", "IFS CODE", "IFSCCODE")):
            continue
        compact_line = re.sub(r"[^A-Z0-9]", "", line)
        for match in re.finditer(rf"(?=({IFSC_CORE_REGEX.pattern}))", compact_line):
            candidate = normalize_ifsc_candidate(match.group(1))
            if is_plausible_ifsc_code(candidate):
                context_score = 15 if any(token in line for token in ("IFSC", "IFS CODE", "IFSCCODE", "NEFT", "RTGS")) else 0
                if context_score <= 0:
                    continue
                candidates.append((context_score, line_index, candidate))
    if not candidates:
        return ""
    candidates.sort(key=lambda item: (item[0], -item[1]), reverse=True)
    return candidates[0][2]


def extract_ifsc_candidates_from_text(text: str) -> list[str]:
    upper = str(text or "").upper()
    candidates: list[str] = []

    label_pattern = re.compile(
        r"\b(?:IFSC|IFS)\s*(?:CODE)?\s*[:/#\-.]?\s*([A-Z0-9][A-Z0-9\s\-/.]{8,24})",
        re.IGNORECASE,
    )
    for match in label_pattern.finditer(upper):
        compact = re.sub(r"[^A-Z0-9]", "", match.group(1))
        if len(compact) >= 11:
            candidates.append(compact[:11])

    compact_upper = re.sub(r"[^A-Z0-9]", "", upper)
    for match in re.finditer(rf"(?=({IFSC_CORE_REGEX.pattern}))", compact_upper):
        candidates.append(match.group(1))

    return [normalize_ifsc_candidate(candidate) for candidate in candidates]


def normalize_ifsc_candidate(value: str) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    if len(raw) != 11:
        return raw
    chars = list(raw)
    for index in (0, 1, 2, 3):
        chars[index] = ocr_char_to_letter(chars[index])
    chars[4] = ocr_char_to_digit(chars[4])
    return "".join(chars)


def is_plausible_ifsc_code(value: str) -> bool:
    candidate = normalize_ifsc_candidate(value)
    if not IFSC_REGEX.fullmatch(candidate):
        return False
    branch_code = candidate[5:]
    if sum(1 for ch in branch_code if ch.isdigit()) < 2:
        return False
    return True


def ifsc_context_keywords() -> tuple[str, ...]:
    return (
        "IFSC",
        "IFSCCODE",
        "RTGS",
        "NEFT",
        "MICR",
        "BANK",
        "BRANCH",
    )


def find_swift_code(text: str) -> str:
    lines = get_clean_lines(str(text or "").upper())
    for index, line in enumerate(lines):
        if not has_swift_label_text(line):
            continue
        candidate_text = line
        if index + 1 < len(lines):
            candidate_text = f"{candidate_text} {lines[index + 1]}"
        for candidate in extract_swift_candidates_from_text(candidate_text):
            if is_plausible_swift_code(candidate):
                return candidate

    candidates: list[tuple[int, int, str]] = []
    for line_index, line in enumerate(lines):
        if not any(token in line for token in ("SWIFT", "BIC", "WIRE", "REMITTANCE")):
            continue
        for candidate in extract_swift_candidates_from_text(line):
            if is_plausible_swift_code(candidate):
                score = 15 if has_swift_label_text(line) else 5
                candidates.append((score, -line_index, candidate))
    if not candidates:
        return ""
    candidates.sort(reverse=True)
    return candidates[0][2]


def has_swift_label_text(value: str) -> bool:
    compact = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    return any(token in compact for token in ("SWIFTCODE", "SWIFTBIC", "BICCODE", "BICSWIFT", "SWIFT", "BIC"))


def extract_swift_candidates_from_text(text: str) -> list[str]:
    upper = str(text or "").upper()
    candidates: list[str] = []
    label_pattern = re.compile(
        r"\b(?:SWIFT(?:\s*/\s*BIC)?|BIC(?:\s*/\s*SWIFT)?)\s*(?:CODE)?\s*[:/#\-.]?\s*([A-Z0-9][A-Z0-9\s\-/.]{7,18})",
        re.IGNORECASE,
    )
    for match in label_pattern.finditer(upper):
        compact = re.sub(r"[^A-Z0-9]", "", match.group(1))
        for length in (11, 8):
            if len(compact) >= length:
                candidates.append(compact[:length])

    for match in SWIFT_REGEX.finditer(upper):
        candidates.append(match.group(0))
    return [normalize_swift_candidate(candidate) for candidate in candidates]


def normalize_swift_candidate(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def is_plausible_swift_code(value: str) -> bool:
    candidate = normalize_swift_candidate(value)
    if not SWIFT_REGEX.fullmatch(candidate):
        return False
    if IFSC_REGEX.fullmatch(candidate):
        return False
    if candidate[:4] in {"BANK", "NAME", "ACCO", "BENE", "BRAN", "SWIF", "CODE"}:
        return False
    return True


def find_iban(text: str) -> str:
    upper = str(text or "").upper()
    lines = get_clean_lines(upper)
    for index, line in enumerate(lines):
        if "IBAN" not in line and "INTERNATIONAL BANK ACCOUNT" not in line:
            continue
        candidate_text = line
        if index + 1 < len(lines):
            candidate_text = f"{candidate_text} {lines[index + 1]}"
        for candidate in extract_iban_candidates_from_text(candidate_text):
            if is_valid_iban(candidate):
                return candidate

    candidates: list[tuple[int, int, str]] = []
    for index, line in enumerate(lines):
        nearby = " ".join(lines[max(0, index - 1) : index + 2])
        nearby_lower = nearby.lower()
        if not has_bank_account_label_text(nearby) and "cont nr" not in nearby_lower:
            continue
        for candidate in extract_iban_candidates_from_text(line):
            if is_valid_iban(candidate):
                score = 20 if has_bank_account_label_text(line) or "cont nr" in line.lower() else 5
                candidates.append((score, -index, candidate))
    if not candidates:
        return ""
    candidates.sort(reverse=True)
    return candidates[0][2]
    return ""


def extract_iban_candidates_from_text(text: str) -> list[str]:
    upper = str(text or "").upper()
    candidates: list[str] = []
    label_pattern = re.compile(
        r"\b(?:IBAN|INTERNATIONAL\s+BANK\s+ACCOUNT\s+(?:NO|NUMBER)?)\s*[:/#\-.]?\s*([A-Z]{2}[0-9O]{2}[A-Z0-9\s\-]{11,42})",
        re.IGNORECASE,
    )
    for match in label_pattern.finditer(upper):
        compact = normalize_iban_candidate(match.group(1))
        for length in range(min(34, len(compact)), 14, -1):
            candidates.append(compact[:length])

    compact_upper = re.sub(r"[^A-Z0-9]", "", upper)
    for match in re.finditer(rf"(?=({IBAN_CORE_REGEX.pattern}))", compact_upper):
        compact = normalize_iban_candidate(match.group(1))
        for length in range(min(34, len(compact)), 14, -1):
            candidates.append(compact[:length])
    return candidates


def normalize_iban_candidate(value: str) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    if len(raw) >= 4:
        chars = list(raw)
        chars[0] = ocr_char_to_letter(chars[0])
        chars[1] = ocr_char_to_letter(chars[1])
        chars[2] = ocr_char_to_digit(chars[2])
        chars[3] = ocr_char_to_digit(chars[3])
        raw = "".join(chars)
    return raw


def is_valid_iban(value: str) -> bool:
    iban = normalize_iban_candidate(value)
    if len(iban) < 15 or len(iban) > 34:
        return False
    if not re.fullmatch(r"[A-Z]{2}[0-9]{2}[A-Z0-9]+", iban):
        return False
    expected_length = IBAN_LENGTH_BY_COUNTRY.get(iban[:2])
    if not expected_length or len(iban) != expected_length:
        return False
    rearranged = iban[4:] + iban[:4]
    converted_parts: list[str] = []
    for ch in rearranged:
        if ch.isdigit():
            converted_parts.append(ch)
        elif "A" <= ch <= "Z":
            converted_parts.append(str(ord(ch) - 55))
        else:
            return False
    remainder = 0
    for ch in "".join(converted_parts):
        remainder = (remainder * 10 + int(ch)) % 97
    return remainder == 1


def find_bank_account_number(text: str) -> str:
    raw_text = str(text or "")
    if looks_like_pan_document(raw_text) and not has_strong_bank_document_marker(raw_text):
        return ""

    labeled = BANK_ACCOUNT_LABEL_VALUE_REGEX.search(raw_text)
    if labeled:
        candidate = normalize_bank_account_number(labeled.group(1))
        if is_plausible_bank_account_number(candidate):
            return candidate

    alphanumeric_labeled = BANK_ACCOUNT_ALPHANUMERIC_LABEL_VALUE_REGEX.search(raw_text.upper())
    if alphanumeric_labeled:
        candidate = normalize_bank_account_number(alphanumeric_labeled.group(1))
        if is_plausible_bank_account_number(candidate):
            return candidate

    lines = get_clean_lines(raw_text)
    for index, line in enumerate(lines):
        if not BANK_ACCOUNT_LABEL_REGEX.search(line) and not has_bank_account_label_text(line):
            continue
        for candidate_line in [line, *lines[index + 1 : index + 4]]:
            for match in BANK_ACCOUNT_NUMBER_CANDIDATE_REGEX.finditer(candidate_line):
                candidate = normalize_bank_account_number(match.group(0))
                if is_plausible_bank_account_number(candidate):
                    return candidate
            for match in BANK_ACCOUNT_ALPHANUMERIC_CANDIDATE_REGEX.finditer(candidate_line):
                candidate = normalize_bank_account_number(match.group(0))
                if is_plausible_bank_account_number(candidate):
                    return candidate

    if looks_like_bank_document(raw_text) or find_ifsc_code(raw_text) or find_swift_code(raw_text) or find_iban(raw_text):
        return find_likely_bank_account_number(lines)
    return ""


def normalize_bank_account_number(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def is_plausible_bank_account_number(value: str) -> bool:
    candidate = normalize_bank_account_number(value)
    if len(candidate) < 5 or len(candidate) > 34:
        return False
    if candidate.isdigit() and len(candidate) < 8:
        return False
    if len(set(candidate)) == 1:
        return False
    if PAN_REGEX.fullmatch(candidate) or GST_REGEX.fullmatch(candidate):
        return False
    if SWIFT_REGEX.fullmatch(candidate) or IFSC_REGEX.fullmatch(candidate) or is_valid_iban(candidate):
        return False
    return any(ch.isdigit() for ch in candidate)


def has_bank_account_label_text(value: str) -> bool:
    lowered = collapse_spaces(value).lower()
    lowered = re.sub(r"\bpermanent\s+account\s+(?:number|no)\b", " ", lowered)
    lowered = re.sub(r"\bpan\s+(?:number|no)\b", " ", lowered)
    return any(
        token in lowered
        for token in (
            "account no",
            "account number",
            "account num",
            "a/c no",
            "a/c number",
            "a/c num",
            "ac no",
            "ac number",
            "acct no",
            "beneficiary account",
            "bank account",
            "cont nr",
            "cont no",
        )
    )


def find_likely_bank_account_number(lines: list[str]) -> str:
    candidates: list[tuple[int, int, str]] = []
    for index, line in enumerate(lines):
        if should_skip_bank_account_candidate_line(line):
            continue
        for match in BANK_ACCOUNT_NUMBER_CANDIDATE_REGEX.finditer(line):
            candidate = normalize_bank_account_number(match.group(0))
            if not is_plausible_bank_account_number(candidate):
                continue
            score = len(candidate)
            nearby_text = " ".join(lines[max(0, index - 2) : index + 3]).lower()
            if has_bank_account_label_text(nearby_text):
                score += 30
            if any(token in nearby_text for token in ("bank", "branch", "ifsc", "neft", "rtgs", "beneficiary")):
                score += 8
            if any(token in nearby_text for token in ("micr", "cheque no", "check no", "mobile", "phone")):
                score -= 15
            candidates.append((score, -index, candidate))

    if not candidates:
        return ""
    candidates.sort(reverse=True)
    return candidates[0][2]


def should_skip_bank_account_candidate_line(value: str) -> bool:
    lowered = collapse_spaces(value).lower()
    blocked_terms = (
        "ifsc",
        "swift",
        "bic",
        "iban",
        "ifs code",
        "micr",
        "cheque no",
        "check no",
        "cheque number",
        "check number",
        "date",
        "mobile",
        "phone",
        "contact",
        "pin code",
        "pincode",
        "postal",
        "amount",
        "rupees",
        "customer id",
        "cust id",
    )
    return any(term in lowered for term in blocked_terms)


def find_beneficiary_name(lines: list[str]) -> str:
    for index, line in enumerate(lines):
        lowered = line.lower()
        if not has_beneficiary_name_label(line):
            continue
        value = extract_beneficiary_name_value(line)
        if looks_like_beneficiary_name_value(value):
            return normalize_title_value(value)
        for next_line in lines[index + 1 : index + 4]:
            if looks_like_beneficiary_name_value(next_line):
                return normalize_title_value(next_line)
    return ""


def has_beneficiary_name_label(value: str) -> bool:
    lowered = collapse_spaces(value).lower()
    blocked = (
        "beneficiary bank",
        "beneficiary branch",
        "beneficiary address",
        "beneficiary account no",
        "beneficiary account number",
        "beneficiary iban",
        "beneficiary swift",
        "beneficiary bic",
    )
    if any(term in lowered for term in blocked):
        return False
    return any(
        re.search(pattern, lowered)
        for pattern in (
            r"\bbeneficiary\s+name\b",
            r"\bname\s+of\s+(?:the\s+)?beneficiary\b",
            r"\baccount\s+holder\s+name\b",
            r"\baccount\s+name\b",
            r"\bbeneficiary\s+account\s+name\b",
            r"\brecipient\s+name\b",
            r"\bpayee\s+name\b",
            r"\bcustomer\s+name\b",
        )
    )


def extract_beneficiary_name_value(value: str) -> str:
    text = collapse_spaces(value)
    patterns = (
        r"\bbeneficiary\s+account\s+name\b",
        r"\bbeneficiary\s+name\b",
        r"\bname\s+of\s+(?:the\s+)?beneficiary\b",
        r"\baccount\s+holder\s+name\b",
        r"\baccount\s+name\b",
        r"\brecipient\s+name\b",
        r"\bpayee\s+name\b",
        r"\bcustomer\s+name\b",
    )
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return clean_bank_text_value(text[match.end() :])
    return ""


def looks_like_beneficiary_name_value(value: str) -> bool:
    text = clean_bank_text_value(value)
    if len(text) < 3 or len(text) > 100:
        return False
    lowered = text.lower()
    blocked = (
        "bank",
        "branch",
        "account number",
        "account no",
        "iban",
        "swift",
        "bic",
        "routing",
        "sort code",
        "address",
        "date",
    )
    if any(term in lowered for term in blocked):
        return False
    if IFSC_REGEX.search(text.upper()) or SWIFT_REGEX.search(text.upper()) or is_valid_iban(text):
        return False
    if re.fullmatch(r"[0-9\s\-]+", text):
        return False
    return any(ch.isalpha() for ch in text)


def clean_bank_text_value(value: str) -> str:
    text = clean_label_value(value)
    text = re.split(
        r"\b(?:bank\s+name|beneficiary\s+bank|branch|account\s+(?:no|number)|a/c|iban|swift|bic|routing|sort\s+code|address)\b",
        text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]
    return text.strip(" :-/|\\")


def find_bank_name(lines: list[str]) -> str:
    value = find_label_value(
        lines,
        (
            "bank name",
            "name of bank",
            "beneficiary bank",
            "beneficiary's bank",
            "receiving bank",
            "intermediary bank",
            "correspondent bank",
            "banker",
            "bankers",
            "drawee bank",
            "financial institution",
        ),
        forbidden=("account number", "account no", "iban", "swift", "bic", "ifsc", "branch", "cheque", "date", "micr"),
    )
    if value and looks_like_bank_name_value(value):
        return normalize_bank_name_value(value)

    for line in lines:
        if not re.search(r"^\s*bank\s*[:/\-|]", line, flags=re.IGNORECASE):
            continue
        value = clean_label_value(re.sub(r"^\s*bank\s*[:/\-|]?", "", line, flags=re.IGNORECASE))
        if looks_like_bank_name_value(value):
            return normalize_bank_name_value(value)

    for line in lines:
        candidate = clean_label_value(line)
        if looks_like_bank_name_value(candidate):
            return normalize_bank_name_value(candidate)
    return ""


def find_branch_name(lines: list[str]) -> str:
    for line in lines:
        candidate = extract_branch_name_from_line(line)
        if candidate:
            return candidate

    value = find_label_value(
        lines,
        ("branch name", "name of branch", "beneficiary branch", "bank branch", "branch", "branch address"),
        forbidden=("ifsc", "swift", "bic", "iban", "account", "cheque", "date", "bank name"),
    )
    if value and looks_like_branch_name_value(value):
        return normalize_title_value(clean_ocr_name_noise(value))
    return ""


def extract_branch_name_from_line(value: str) -> str:
    text = collapse_spaces(value)
    lowered = text.lower()
    if "branch code" in lowered or "all branches" in lowered:
        return ""
    match = re.search(r"\b([A-Za-z][A-Za-z0-9 .,'/&\-]{2,90})\s+Branch\b", text, flags=re.IGNORECASE)
    if not match:
        return ""
    candidate = match.group(1)
    if "-" in candidate:
        candidate = candidate.rsplit("-", 1)[-1]
    candidate = clean_ocr_name_noise(candidate)
    if not looks_like_branch_name_value(candidate):
        return ""
    return normalize_title_value(candidate)


def find_label_value(
    lines: list[str],
    labels: tuple[str, ...],
    *,
    forbidden: tuple[str, ...] = (),
) -> str:
    normalized_labels = tuple(label.lower() for label in labels)
    forbidden_terms = tuple(term.lower() for term in forbidden)

    for index, line in enumerate(lines):
        lowered = line.lower()
        if any(term in lowered for term in forbidden_terms):
            continue
        for label in normalized_labels:
            if label not in lowered:
                continue
            value = line[lowered.find(label) + len(label) :]
            value = clean_label_value(value)
            if value and looks_like_name_value(value):
                return value
            for next_line in lines[index + 1 : index + 3]:
                if looks_like_name_value(next_line) and not any(term in next_line.lower() for term in forbidden_terms):
                    return clean_label_value(next_line)
    return ""


def find_pan_name(lines: list[str], pan_number: str = "") -> str:
    for index, line in enumerate(lines):
        if is_pan_non_holder_name_line(line):
            continue

        value = extract_inline_pan_name_value(line)
        if looks_like_pan_holder_name(value):
            return clean_pan_holder_name(value)

        if has_pan_name_label(line):
            for next_line in lines[index + 1 : index + 5]:
                candidate = clean_pan_holder_name(next_line)
                if looks_like_pan_holder_name(candidate):
                    return candidate

    for index, line in enumerate(lines):
        lowered = line.lower()
        if "father" not in lowered and "parent" not in lowered:
            continue
        for previous_line in reversed(lines[max(0, index - 5) : index]):
            candidate = clean_pan_holder_name(previous_line)
            if looks_like_pan_holder_name(candidate):
                return candidate

    pan_indexes = [
        index
        for index, line in enumerate(lines)
        if (pan_number and pan_number in re.sub(r"[^A-Z0-9]", "", line.upper())) or PAN_REGEX.search(line.upper())
    ]
    for pan_index in pan_indexes:
        for candidate_line in reversed(lines[max(0, pan_index - 6) : pan_index]):
            candidate = clean_pan_holder_name(candidate_line)
            if looks_like_pan_holder_name(candidate):
                return candidate
        for candidate_line in lines[pan_index + 1 : pan_index + 6]:
            candidate = clean_pan_holder_name(candidate_line)
            if looks_like_pan_holder_name(candidate):
                return candidate

    for index, line in enumerate(lines):
        lowered = line.lower()
        if "father" in lowered or "date of birth" in lowered:
            continue
        if lowered in {"name", "name:"} or re.search(r"\bname\b", lowered):
            value = clean_label_value(re.sub(r"\bname\b", "", line, flags=re.IGNORECASE))
            if looks_like_name_value(value):
                return value
            for next_line in lines[index + 1 : index + 4]:
                lowered_next = next_line.lower()
                if "father" in lowered_next or "date of birth" in lowered_next:
                    continue
                if looks_like_name_value(next_line):
                    return clean_label_value(next_line)
    return ""


def has_pan_name_label(value: str) -> bool:
    lowered = collapse_spaces(value).lower()
    if "father" in lowered or "mother" in lowered or "parent" in lowered:
        return False
    return any(
        re.search(pattern, lowered)
        for pattern in (
            r"\bname\s+as\s+per\s+pan\b",
            r"\bpan\s+(?:card\s+)?holder\s+name\b",
            r"\bcard\s+holder\s+name\b",
            r"\bapplicant\s+name\b",
            r"\bassessee\s+name\b",
            r"\bname\s+of\s+(?:the\s+)?(?:applicant|card\s+holder|assessee)\b",
            r"\bfull\s+name\b",
            r"^\s*name\s*[:/\-|]?\s*$",
            r"^\s*name\s*[:/\-|]",
        )
    )


def extract_inline_pan_name_value(value: str) -> str:
    text = collapse_spaces(value)
    if not text:
        return ""
    patterns = (
        r"\bname\s+as\s+per\s+pan\b",
        r"\bpan\s+(?:card\s+)?holder\s+name\b",
        r"\bcard\s+holder\s+name\b",
        r"\bapplicant\s+name\b",
        r"\bassessee\s+name\b",
        r"\bname\s+of\s+(?:the\s+)?(?:applicant|card\s+holder|assessee)\b",
        r"\bfull\s+name\b",
        r"^\s*name\b",
    )
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue
        return clean_pan_holder_name(text[match.end() :])
    return ""


def clean_pan_holder_name(value: str) -> str:
    text = collapse_spaces(value)
    text = clean_ocr_name_noise(text)
    text = re.sub(r"^[\s:;|/\\.\-]+", "", text)
    text = re.split(
        r"\b(?:father|mother|parent|date\s+of\s+birth|dob|date\s+of\s+incorporation|signature|permanent\s+account|"
        r"account\s+number|pan\s+number|pan\s+no)\b",
        text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]
    text = PAN_REGEX.sub("", text.upper() if text.isupper() else text)
    text = re.sub(r"\b[0-9]{1,2}[\-/\.][0-9]{1,2}[\-/\.][0-9]{2,4}\b", "", text)
    text = clean_label_value(text)
    return text[:100].strip(" :-/|\\")


def clean_ocr_name_noise(value: str) -> str:
    text = collapse_spaces(value)
    if not text:
        return ""
    text = "".join(ch if ch.isascii() else " " for ch in text)
    text = re.sub(r"[^A-Za-z0-9 &.,'()/\\-]+", " ", text)
    text = collapse_spaces(text).strip(" ,.;:-/|\\")
    tokens = text.split()
    suffix_allowlist = {"co", "corp", "inc", "llc", "llp", "ltd", "pvt"}
    uppercase_prefix_count = sum(1 for token in tokens[:-1] if token.strip(".,").isupper() and len(token.strip(".,")) > 1)
    while len(tokens) > 2:
        tail = tokens[-1].strip(".,")
        if tail.lower() in suffix_allowlist:
            break
        if uppercase_prefix_count >= 2 and len(tail) <= 4 and not tail.isupper():
            tokens.pop()
            continue
        break
    return " ".join(tokens)


def is_pan_non_holder_name_line(value: str) -> bool:
    text = collapse_spaces(value)
    if not text:
        return True
    lowered = text.lower()
    if has_pan_name_label(text):
        return False
    blocked_phrases = (
        "income tax department",
        "govt. of india",
        "govt of india",
        "government of india",
        "permanent account number",
        "signature",
        "date of birth",
        "date of incorporation",
        "father",
        "mother",
        "parent",
        "pan card",
        "tax department",
    )
    if any(phrase in lowered for phrase in blocked_phrases):
        return True
    if PAN_REGEX.search(text.upper()):
        return True
    if re.search(r"\b[0-9]{1,2}[\-/\.][0-9]{1,2}[\-/\.][0-9]{2,4}\b", text):
        return True
    letters = sum(1 for ch in text if ch.isalpha())
    digits = sum(1 for ch in text if ch.isdigit())
    if digits and digits >= letters:
        return True
    return False


def looks_like_pan_holder_name(value: str) -> bool:
    text = clean_pan_holder_name(value)
    if not looks_like_name_value(text):
        return False
    if is_pan_non_holder_name_line(text):
        return False
    lowered = text.lower()
    blocked = (
        "permanent",
        "account",
        "department",
        "government",
        "govt",
        "signature",
        "birth",
        "incorporation",
        "number",
        "card",
    )
    return not any(term in lowered for term in blocked)


def clean_label_value(value: str) -> str:
    text = collapse_spaces(value)
    text = re.sub(r"^[\s:;|/\\.\-]+", "", text)
    text = re.sub(r"\s*(?:\||:)\s*$", "", text)
    text = re.sub(r"\b(GSTIN|PAN|CIN|UDYAM|ESIC?|PF)\b.*$", "", text, flags=re.IGNORECASE).strip()
    return text[:100].strip()


def clean_identifier_value(value: str, *, allowed: str = "") -> str:
    chars = []
    allowed_set = set(allowed)
    for ch in str(value or "").upper():
        if ch.isalnum() or ch in allowed_set:
            chars.append(ch)
        elif ch.isspace() and "/" in allowed_set:
            chars.append("/")
    return re.sub(r"/+", "/", "".join(chars)).strip("/- ")


def trim_identifier(value: str) -> str:
    text = str(value or "").strip().upper()
    text = re.split(r"\b(?:NAME|ADDRESS|DATE|STATUS|VALID|CERTIFICATE)\b", text, maxsplit=1)[0]
    return text.strip("/-: ")


def looks_like_name_value(value: str) -> bool:
    text = clean_label_value(value)
    if len(text) < 3 or len(text) > 100:
        return False
    lowered = text.lower()
    blocked = (
        "government",
        "income tax",
        "department",
        "certificate",
        "registration",
        "address",
        "date",
        "number",
        "status",
        "signature",
        "father",
    )
    if any(term in lowered for term in blocked):
        return False
    return any(ch.isalpha() for ch in text)


def looks_like_bank_name_value(value: str) -> bool:
    text = clean_label_value(value)
    if len(text) < 3 or len(text) > 100:
        return False
    lowered = text.lower()
    blocked = (
        "account",
        "branch",
        "ifsc",
        "swift",
        "bic",
        "iban",
        "micr",
        "cheque",
        "check",
        "cancelled",
        "date",
        "signature",
        "pay",
        "rupees",
    )
    if any(term in lowered for term in blocked):
        return False
    if IFSC_REGEX.search(text.upper()) or re.search(r"\b[0-9]{6,}\b", text):
        return False
    bank_markers = (
        "bank",
        "hdfc",
        "icici",
        "axis",
        "sbi",
        "kotak",
        "indusind",
        "idfc",
        "idbi",
        "pnb",
        "punjab national",
        "canara",
        "union bank",
        "bank of",
        "central bank",
        "indian bank",
        "federal bank",
        "rbl",
        "bandhan",
        "yes bank",
        "au small finance",
        "city union",
        "karur",
        "south indian",
        "standard chartered",
        "hsbc",
        "dbs",
        "jpmorgan",
        "jp morgan",
        "chase",
        "citibank",
        "citi bank",
        "citi",
        "bank of america",
        "wells fargo",
        "deutsche",
        "barclays",
        "lloyds",
        "natwest",
        "royal bank",
        "scotiabank",
        "td bank",
        "rbc",
        "ubs",
        "credit suisse",
        "bnp paribas",
        "societe generale",
        "santander",
        "ing bank",
        "rabobank",
        "commerzbank",
        "unicredit",
        "mizuho",
        "sumitomo",
        "mitsubishi",
        "dbs bank",
        "ocbc",
        "uob",
        "emirates nbd",
        "qatar national",
        "doha bank",
        "first abu dhabi",
        "abu dhabi commercial",
        "bank of china",
        "china construction",
        "industrial and commercial bank",
        "anz",
        "westpac",
        "nab",
        "commonwealth bank",
    )
    return any(marker in lowered for marker in bank_markers)


def looks_like_branch_name_value(value: str) -> bool:
    text = clean_ocr_name_noise(clean_label_value(value))
    if len(text) < 2 or len(text) > 100:
        return False
    lowered = text.lower()
    if lowered in {"code", "branch code", "codul sucursalei"}:
        return False
    blocked = (
        "account",
        "ifsc",
        "swift",
        "bic",
        "iban",
        "micr",
        "cheque",
        "check",
        "cancelled",
        "date",
        "signature",
        "pay",
        "rupees",
    )
    if any(term in lowered for term in blocked):
        return False
    if IFSC_REGEX.search(text.upper()) or re.search(r"\b[0-9]{6,}\b", text):
        return False
    return sum(1 for ch in text if ch.isalpha()) >= 3


def normalize_title_value(value: str) -> str:
    text = clean_label_value(value)
    if text.isupper() or text.islower():
        return text.title()
    return text


def normalize_bank_name_value(value: str) -> str:
    raw_text = collapse_spaces(value)
    lowered = raw_text.lower()
    for marker, label in KNOWN_BANK_NAME_NORMALIZATIONS:
        if marker in lowered:
            return label
    text = clean_ocr_name_noise(clean_label_value(value))
    text = text.replace("]", "I").replace("|", "I").replace("[", "I")
    return text.title()


def derive_business_entity_from_pan(pan: str) -> str:
    normalized = str(pan or "").strip().upper()
    if len(normalized) < 4:
        return ""
    return PAN_ENTITY_BY_FOURTH_CHAR.get(normalized[3], "")


def find_msme_category(text: str) -> str:
    lower = str(text or "").lower()
    for token, label in MSME_CATEGORY_LABELS.items():
        if token in lower:
            return label
    return ""


def find_msme_industry(text: str) -> str:
    lower = str(text or "").lower()
    if "manufactur" in lower:
        return "Manufacturing"
    for token, label in MSME_INDUSTRY_LABELS.items():
        if token in lower:
            return label
    return ""


def looks_like_msme_document(text: str) -> bool:
    lower = str(text or "").lower()
    return any(term in lower for term in ("udyam", "msme", "micro small and medium", "enterprise type"))


def looks_like_gst_document(text: str) -> bool:
    lower = str(text or "").lower()
    return any(term in lower for term in ("gstin", "goods and services tax", "gst registration", "taxpayer type"))


def looks_like_pan_document(text: str) -> bool:
    lower = str(text or "").lower()
    return any(term in lower for term in ("permanent account number", "income tax department", "pan card"))


def looks_like_bank_document(text: str) -> bool:
    return has_strong_bank_document_marker(text)


def has_strong_bank_document_marker(text: str) -> bool:
    lower = collapse_spaces(text).lower()
    if not lower:
        return False
    pan_safe_lower = re.sub(r"\bpermanent\s+account\s+(?:number|no)\b", " ", lower)
    pan_safe_lower = re.sub(r"\bpan\s+(?:number|no)\b", " ", pan_safe_lower)
    strong_terms = (
        "cancelled cheque",
        "cancelled check",
        "bank certificate",
        "banker's certificate",
        "ifsc",
        "ifs code",
        "swift",
        "swift code",
        "bic",
        "bic code",
        "iban",
        "international bank account",
        "beneficiary account",
        "beneficiary name",
        "beneficiary bank",
        "account holder",
        "bank account",
        "bank name",
        "branch name",
        "wire transfer",
        "wire instructions",
        "remittance",
        "routing number",
        "sort code",
        "cheque",
        "check",
        "micr",
        "neft",
        "rtgs",
        "current account",
        "savings account",
    )
    if any(term in pan_safe_lower for term in strong_terms):
        return True
    if "bank" in pan_safe_lower and any(
        term in pan_safe_lower for term in ("account", "branch", "beneficiary", "current", "savings", "statement")
    ):
        return True
    return has_bank_account_label_text(pan_safe_lower) and not has_pan_account_label_text(lower)


def has_pan_account_label_text(value: str) -> bool:
    lowered = collapse_spaces(value).lower()
    return any(
        token in lowered
        for token in (
            "permanent account number",
            "permanent account no",
            "pan number",
            "pan no",
        )
    )


def detect_document_types(text: str, fields: dict[str, str], filename: str = "") -> list[str]:
    combined = f"{filename}\n{text}".lower()
    detected: list[str] = []

    def add(document_type: str) -> None:
        if document_type not in detected:
            detected.append(document_type)

    if fields.get("gstNumber") or any(term in combined for term in ("gstin", "goods and services tax", "gst registration")):
        add("gst")
    if fields.get("udyamNumber") or any(term in combined for term in ("udyam", "msme", "micro small and medium")):
        add("msme")
    if fields.get("cinNumber") or any(term in combined for term in ("certificate of incorporation", "corporate identity", "ministry of corporate affairs")):
        add("cin")
    if any(term in combined for term in ("permanent account number", "income tax department", "pan card")):
        add("pan")
    if fields.get("pfNumber") or any(term in combined for term in ("provident fund", "epfo", "employee provident")):
        add("pf")
    if fields.get("esiNumber") or any(term in combined for term in ("employees' state insurance", "employee state insurance", "esic")):
        add("esi")
    if (
        fields.get("ifscCode")
        or fields.get("swiftCode")
        or fields.get("iban")
        or fields.get("bankAccountNumber")
        or looks_like_bank_document(combined)
    ):
        add("bank")
    return detected


def determine_expected_document_match(
    expected_type: str,
    detected_types: list[str],
    fields: dict[str, str],
) -> bool | None:
    if not expected_type:
        return None
    if expected_type in detected_types:
        return True
    expected_key_map = {
        "gst": ("gstNumber",),
        "pan": ("panNumber",),
        "msme": ("udyamNumber", "msmeCategory", "msmeIndustry"),
        "pf": ("pfNumber",),
        "cin": ("cinNumber",),
        "esi": ("esiNumber",),
        "bank": (
            "ifscCode",
            "swiftCode",
            "iban",
            "bankAccountNumber",
            "bankName",
            "branchName",
            "beneficiaryName",
        ),
    }
    if any(fields.get(key) for key in expected_key_map.get(expected_type, ())):
        return True
    if detected_types:
        return False
    return False


def build_form_field_aliases(fields: dict[str, str]) -> dict[str, str]:
    output: dict[str, str] = {}

    def add(key: str, value: Any) -> None:
        text = str(value or "").strip()
        if text:
            output[key] = text

    gst_number = fields.get("gstNumber", "")
    add("gstNumber", gst_number)
    add("GSTIN", gst_number)

    pan_number = fields.get("panNumber", "")
    add("panNumber", pan_number)
    add("PANNumber", pan_number)
    add("PANNo", pan_number)

    add("cinNumber", fields.get("cinNumber"))
    add("CIN", fields.get("cinNumber"))

    add("udyamNumber", fields.get("udyamNumber"))
    add("UdyamNumber", fields.get("udyamNumber"))

    add("pfNumber", fields.get("pfNumber"))
    add("PFNumber", fields.get("pfNumber"))

    add("esiNumber", fields.get("esiNumber"))
    add("ESIN", fields.get("esiNumber"))

    add("ifscCode", fields.get("ifscCode"))
    add("IFSC", fields.get("ifscCode"))
    add("IFSC_Code", fields.get("ifscCode"))

    add("swiftCode", fields.get("swiftCode"))
    add("SwiftCode", fields.get("swiftCode"))
    add("SWIFT", fields.get("swiftCode"))
    add("SWIFTCode", fields.get("swiftCode"))
    add("BIC", fields.get("swiftCode"))

    add("iban", fields.get("iban"))
    add("IBAN", fields.get("iban"))

    add("beneficiaryName", fields.get("beneficiaryName"))
    add("BeneficiaryName", fields.get("beneficiaryName"))
    add("AccountHolderName", fields.get("beneficiaryName"))

    add("bankName", fields.get("bankName"))
    add("BankName", fields.get("bankName"))

    add("branchName", fields.get("branchName"))
    add("BranchName", fields.get("branchName"))

    add("bankAccountNumber", fields.get("bankAccountNumber"))
    add("BankAccount", fields.get("bankAccountNumber"))
    add("BankAccountNumber", fields.get("bankAccountNumber"))

    add("tradeNameGst", fields.get("tradeNameGst"))
    add("TradeNameGST", fields.get("tradeNameGst"))
    add("TradeName", fields.get("tradeNameGst"))

    add("taxpayerTypeGst", fields.get("taxpayerTypeGst"))
    add("TaxpayerTypeGST", fields.get("taxpayerTypeGst"))

    add("tradeNamePan", fields.get("tradeNamePan"))
    add("TradeNamePAN", fields.get("tradeNamePan"))

    add("businessEntityType", fields.get("businessEntityType"))
    add("TypeofEntity", fields.get("businessEntityType"))
    add("BusinessEntityType", fields.get("businessEntityType"))

    add("panDate", fields.get("panDate"))
    add("PANDate", fields.get("panDate"))
    add("dateOfBirth", fields.get("dateOfBirth"))
    add("DateOfBirth", fields.get("dateOfBirth"))
    add("DOB", fields.get("dateOfBirth"))
    add("dateOfIncorporation", fields.get("dateOfIncorporation"))
    add("DateOfIncorporation", fields.get("dateOfIncorporation"))
    add("chequeDate", fields.get("chequeDate"))
    add("ChequeDate", fields.get("chequeDate"))
    add("bankDocumentDate", fields.get("bankDocumentDate"))
    add("BankDocumentDate", fields.get("bankDocumentDate"))

    add("msmeCategory", fields.get("msmeCategory"))
    add("MSMECategory", fields.get("msmeCategory"))
    add("msmeIndustry", fields.get("msmeIndustry"))
    add("MSMEIndustry", fields.get("msmeIndustry"))

    return output


def score_confidence(
    expected_type: str,
    detected_types: list[str],
    fields: dict[str, str],
    has_text: bool,
) -> float:
    score = 0.15 if has_text else 0.0
    if expected_type and expected_type in detected_types:
        score += 0.35
    elif detected_types:
        score += 0.2
    score += min(0.5, len([value for value in fields.values() if value]) * 0.1)
    return round(min(score, 1.0), 2)
