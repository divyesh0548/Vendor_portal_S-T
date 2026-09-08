import base64
import hashlib
import itertools
import json
import mimetypes
import os
import re
import shutil
import smtplib
import subprocess
import tempfile
import time
import threading
import uuid
import zipfile
import importlib.util
from io import BytesIO
from contextlib import asynccontextmanager
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from html import escape as html_escape
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote, unquote, urlparse
from urllib.request import Request as UrlRequest, urlopen

import jwt
from docx import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph as DocxParagraph
from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException, Request, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from starlette.datastructures import UploadFile as StarletteUploadFile

from .document_extraction import register_document_extraction_routes
from .db import (
    authenticate_db_user,
    cleanup_vendor_duplicate_columns,
    get_all_vendor_records,
    get_db_status,
    get_vendor_duplicate_records_by_pan,
    get_vendor_record_by_record_id,
    get_vendor_document_content_by_record_id,
    get_vendor_document_content_by_filename,
    get_vendor_records_by_assignee_email,
    get_vendor_records_by_approver_status,
    get_db_users,
    get_db_users_by_role,
    initialize_database,
    insert_db_user,
    update_db_user_by_email,
    ensure_vendor_columns,
    update_db_user_password_by_email,
    insert_vendor_master_data,
    update_vendor_approver_status_by_record_id,
    update_vendor_record_by_record_id,
)
from .sap_process import kill_excel_processes, kill_sap_processes, run_post_vendor_upload_process


JWT_SECRET = os.getenv("JWT_SECRET", "dev-secret-change-me")
JWT_ALGORITHM = "HS256"
SEED_SAMPLE_USERS = os.getenv("SEED_SAMPLE_USERS", "true").lower() == "true"
AUTH_COOKIE_NAME = "auth_token"
DEFAULT_NEW_USER_PASSWORD = os.getenv("DEFAULT_NEW_USER_PASSWORD", "R_Vendor@1234")
WEB_PUBLIC_PATHS = {
    "/web/login",
    "/web/code-of-conduct-docx",
    "/web/code-of-conduct-pdf",
    "/web/code-of-conduct-preview",
    "/web/code-of-conduct-viewer",
    "/web/logo-image",
    "/web/header-art-image",
    "/web/portal-liquid.css",
    "/web/customer-public-form",
    "/web/update-customer",
}
PUBLIC_VENDOR_WEB_PATHS = {
    "/web/vendor-registration",
    "/web/vendor-registration/prescreen",
    "/web/vendor-registration/details",
    "/web/update-vendor",
}


def _normalize_smtp_password(value: str) -> str:
    raw = str(value or "").strip()
    if len(raw) >= 2 and ((raw[0] == raw[-1] == '"') or (raw[0] == raw[-1] == "'")):
        raw = raw[1:-1].strip()
    return "".join(raw.split())


def _first_non_empty_env(*names: str) -> str:
    for name in names:
        value = os.getenv(str(name or "").strip(), "")
        text = str(value or "").strip()
        if text:
            return text
    return ""


SMTP_HOST = _first_non_empty_env("SMTP_HOST", "OUTLOOK_SMTP_HOST") or "smtp.gmail.com"
try:
    SMTP_PORT = int((_first_non_empty_env("SMTP_PORT", "OUTLOOK_SMTP_PORT") or "587").strip())
except (TypeError, ValueError):
    SMTP_PORT = 587
SMTP_SENDER_EMAIL = _first_non_empty_env("SMTP_SENDER_EMAIL", "OUTLOOK_SENDER_EMAIL")
SMTP_USERNAME = _first_non_empty_env("SMTP_USERNAME", "OUTLOOK_SMTP_USERNAME") or SMTP_SENDER_EMAIL
SMTP_PASSWORD = _normalize_smtp_password(
    _first_non_empty_env(
        "SMTP_PASSWORD",
        "OUTLOOK_SMTP_PASSWORD",
        "GMAIL_APP_PASSWORD",
        "GOOGLE_APP_PASSWORD",
        "MAIL_APP_PASSWORD",
    )
)
SMTP_USE_STARTTLS = os.getenv("SMTP_USE_STARTTLS", os.getenv("OUTLOOK_USE_STARTTLS", "true")).strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "").strip().rstrip("/")
try:
    SMTP_TIMEOUT_SECONDS = int(
        os.getenv("SMTP_TIMEOUT_SECONDS", os.getenv("OUTLOOK_SMTP_TIMEOUT_SECONDS", "30")).strip()
    )
except (TypeError, ValueError):
    SMTP_TIMEOUT_SECONDS = 30


@asynccontextmanager
async def lifespan(_: FastAPI):
    seed_users()
    initialize_database()
    yield


app = FastAPI(title="Vendor Backend (Python, No SQL)", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

register_document_extraction_routes(app)


security = HTTPBearer(auto_error=False)
user_id_counter = itertools.count(1)
form_id_counter = itertools.count(1)
approval_id_counter = itertools.count(1)

users: dict[int, dict[str, Any]] = {}
users_by_email: dict[str, int] = {}
vendor_forms: dict[int, dict[str, Any]] = {}
invite_forms: dict[str, dict[str, Any]] = {}
approval_sequence: dict[str, Any] | None = None
mock_files: dict[str, dict[str, Any]] = {}
UPLOADED_FILES_DIR = Path(__file__).parent / "uploaded_files"
UPLOADED_FILES_DIR.mkdir(parents=True, exist_ok=True)
SAP_VENDOR_TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "Vendor Template" / "Vendor Template.xlsx"
SAP_CONFIG_WORKBOOK_PATH = Path(__file__).resolve().parents[1] / "Vendor Template" / "Config.xlsx"
CODE_OF_CONDUCT_TEMPLATE_PATH = (
    Path(__file__).resolve().parents[1]
    / "Data Validation"
    / "Supplier_and_Contractor_Code_of_Conduct_Sharp_Tannan.docx"
)
CODE_OF_CONDUCT_DOCUMENT_FIELD = "codeOfConductDocument"
CODE_OF_CONDUCT_RENDER_LOCK = threading.Lock()
SAP_CONFIG_INPUT_SHEET_NAME = "Input File"
SAP_VENDOR_UPLOAD_PATH = Path(
    os.getenv(
        "SAP_VENDOR_UPLOAD_PATH",
        r"D:\Vendor Upload Template - BOT.xlsx",
    )
)
SAP_LOGON_EXECUTABLE_PATH = Path(
    os.getenv(
        "SAP_LOGON_EXECUTABLE_PATH",
        r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
    )
)
SAP_OPEN_VENDOR_UPLOAD_AFTER_SYNC = os.getenv("SAP_OPEN_VENDOR_UPLOAD_AFTER_SYNC", "false").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}
SAP_CLOSE_VENDOR_UPLOAD_AFTER_SYNC = os.getenv("SAP_CLOSE_VENDOR_UPLOAD_AFTER_SYNC", "true").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}
SAP_AUTO_LAUNCH_LOGON = os.getenv("SAP_AUTO_LAUNCH_LOGON", "true").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}
SAP_AUTO_LOGON_AFTER_LAUNCH = os.getenv("SAP_AUTO_LOGON_AFTER_LAUNCH", "true").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}
SAP_SERVER_NAME = os.getenv("SAP_SERVER_NAME", "DEV Server").strip() or "DEV Server"
SAP_LOGON_USERNAME = os.getenv("SAP_LOGON_USERNAME", "").strip()
SAP_LOGON_PASSWORD = os.getenv("SAP_LOGON_PASSWORD", "")
SAP_LOGON_CLIENT = os.getenv("SAP_LOGON_CLIENT", "400").strip() or "400"
SAP_LOGON_LANGUAGE = os.getenv("SAP_LOGON_LANGUAGE", "EN").strip() or "EN"
SAP_VENDOR_UPLOAD_TRANSACTION_CODE = os.getenv("SAP_VENDOR_UPLOAD_TRANSACTION_CODE", "ZVEND_UP").strip() or "ZVEND_UP"
try:
    SAP_GUI_WAIT_TIMEOUT_SECONDS = int(os.getenv("SAP_GUI_WAIT_TIMEOUT_SECONDS", "120").strip())
except (TypeError, ValueError):
    SAP_GUI_WAIT_TIMEOUT_SECONDS = 120
try:
    SAP_TRANSACTION_WAIT_SECONDS = int(os.getenv("SAP_TRANSACTION_WAIT_SECONDS", "20").strip())
except (TypeError, ValueError):
    SAP_TRANSACTION_WAIT_SECONDS = 20
try:
    SAP_EXCEL_MINIMIZE_DELAY_SECONDS = int(os.getenv("SAP_EXCEL_MINIMIZE_DELAY_SECONDS", "3").strip())
except (TypeError, ValueError):
    SAP_EXCEL_MINIMIZE_DELAY_SECONDS = 3
SAP_QUERY_REFRESH_TIMEOUT_SECONDS = 300
SAP_VENDOR_TEMPLATE_HEADERS = [
    "Vendor_Type",
    "VendorName1",
    "VendorName2",
    "GST_Number",
    "PAN_Number",
    "PANName",
    "UdyamNumber",
    "MSMECategory",
    "MSMEIndustry",
    "VendorCIN",
    "Address1",
    "Address2",
    "Address3",
    "PostalCode",
    "City",
    "Country",
    "BankName",
    "Branch",
    "BankAccount",
    "IFSC",
    "IncoTerms",
    "PaymentTerms",
    "Mobile",
    "TaxCode",
    "VendorEmail",
    "CompanyDealing",
    "StateName",
    "TurnoverLimit",
    "MSMEDate",
]
SAP_SYNC_LOCK = threading.Lock()
STATUS_PENDING_SAP_CODE_CREATION = "Pending for SAP Code Creation"


def _normalize_workflow_status_key(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"\s+", " ", text)
    replacements = {
        "pendign": "pending",
        "creattion": "creation",
        "cretion": "creation",
    }
    for old_value, new_value in replacements.items():
        text = text.replace(old_value, new_value)
    return text.strip()


def _is_pending_sap_code_creation_status(value: Any) -> bool:
    return _normalize_workflow_status_key(value) == "pending for sap code creation"


def _updates_request_sap_code_creation(updates: dict[str, Any] | None) -> bool:
    if not isinstance(updates, dict):
        return False

    status_keys = {
        "approverstatus",
        "appstatus",
        "approvstatus",
        "invitestatus",
        "status",
    }
    for key, value in updates.items():
        normalized_key = str(key or "").strip().lower().replace(" ", "")
        if normalized_key in status_keys and _is_pending_sap_code_creation_status(value):
            return True
    return False


DOCUMENT_FIELDS = {
    "gstDocument",
    "panDocument",
    "cinDocument",
    "msmeDocument",
    "pfDocument",
    "esiDocument",
    "cancelledCheque",
    "inhouseMachineriesDoc",
    "numberOfPlantsDoc",
    "iso1",
    "iso2",
    "iso3",
    "iso4",
    "iso5",
    "ecoVadisDoc",
    "ohsasDoc",
    "relevantCertificateDoc",
    CODE_OF_CONDUCT_DOCUMENT_FIELD,
}

DOCUMENT_FIELD_ALIASES: dict[str, list[str]] = {
    "gstDocument": ["GSTAttachment", "GSTCertificateName"],
    "panDocument": ["PANAttachment", "PAN_Card_Filename"],
    "cinDocument": ["CINCerti", "Incorporation_Filename"],
    "msmeDocument": ["MSMECerti", "MSME_Certificate_Filename"],
    "pfDocument": ["PFCerti", "PF_Certificate_Filename"],
    "esiDocument": ["ESICerti", "ESI_Filename"],
    "cancelledCheque": ["CancelledCheque", "Cancelled_Cheque_Filename"],
    "inhouseMachineriesDoc": ["InhouseMachineAttchment", "Machineries_List_Filename"],
    "numberOfPlantsDoc": ["TopCustomerAttachment"],
    "iso1": ["ISO_File1_Name"],
    "iso2": ["ISO_File2_Name"],
    "iso3": ["ISO_File3_Name"],
    "iso4": ["ISO_File4_Name"],
    "iso5": ["ISO_File5_Name"],
    "ecoVadisDoc": ["ECOVadis_Filename"],
    "ohsasDoc": ["OHSAS_Filename"],
    "relevantCertificateDoc": ["Other_Filename"],
    CODE_OF_CONDUCT_DOCUMENT_FIELD: ["CodeOfConductDocument"],
}


def _normalize_mock_file_reference(file_name: str) -> tuple[str, str]:
    cleaned_parts: list[str] = []
    decoded_name = unquote(str(file_name or ""))
    for raw_part in decoded_name.replace("\\", "/").split("/"):
        part = Path(raw_part).name.strip()
        if not part or part in {".", ".."}:
            continue
        cleaned_parts.append(part)

    if not cleaned_parts:
        return "", ""

    storage_key = "/".join(cleaned_parts)
    return storage_key, cleaned_parts[-1]


def _resolve_uploaded_file_disk_path(storage_key: str) -> Path | None:
    normalized_key, _ = _normalize_mock_file_reference(storage_key)
    if not normalized_key:
        return None

    candidate_path = UPLOADED_FILES_DIR.joinpath(*normalized_key.split("/"))
    try:
        resolved_candidate = candidate_path.resolve()
        resolved_base = UPLOADED_FILES_DIR.resolve()
        resolved_candidate.relative_to(resolved_base)
    except Exception:
        return None
    return resolved_candidate


def _sanitize_uploaded_file_segment(value: str, default: str) -> str:
    sanitized = re.sub(r"[^A-Za-z0-9._-]+", "-", str(value or "").strip()).strip("._-")
    return sanitized or default


def _uploaded_document_storage_dir(record_id: str, field_name: str) -> Path:
    record_segment = _sanitize_uploaded_file_segment(record_id, "record")
    field_segment = _sanitize_uploaded_file_segment(field_name, "document")
    return UPLOADED_FILES_DIR / record_segment / field_segment


def _pick_latest_file(paths: list[Path]) -> Path | None:
    existing_files = [path for path in paths if path.exists() and path.is_file()]
    if not existing_files:
        return None
    existing_files.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    return existing_files[0]


def _find_uploaded_file_by_name(file_name: str) -> Path | None:
    normalized_name = Path(str(file_name or "")).name.strip()
    if not normalized_name or not UPLOADED_FILES_DIR.exists():
        return None

    try:
        return _pick_latest_file(list(UPLOADED_FILES_DIR.rglob(normalized_name)))
    except Exception:
        return None


def _find_uploaded_file_for_record_field(record_id: str, field_name: str, file_name: str = "") -> Path | None:
    storage_dir = _uploaded_document_storage_dir(record_id, field_name)
    if not storage_dir.exists() or not storage_dir.is_dir():
        return None

    normalized_name = Path(str(file_name or "")).name.strip()
    if normalized_name:
        direct_match = storage_dir / normalized_name
        if direct_match.exists() and direct_match.is_file():
            return direct_match

    return _pick_latest_file(list(storage_dir.iterdir()))


def _response_from_uploaded_disk_file(file_path: Path) -> FileResponse:
    guessed_content_type = mimetypes.guess_type(str(file_path))[0]
    header_bytes = b""
    try:
        with file_path.open("rb") as file_handle:
            header_bytes = file_handle.read(4096)
    except Exception:
        header_bytes = b""
    sniffed_content_type = _sniff_document_content_type(header_bytes)
    media_type = sniffed_content_type or guessed_content_type or "application/octet-stream"
    headers = {"Content-Disposition": f'inline; filename="{file_path.name}"'}
    return FileResponse(file_path, media_type=media_type, headers=headers)


def _normalize_inline_document_content_type(file_name: str, content_type: Any) -> str:
    raw_content_type = str(content_type or "").split(";", 1)[0].strip().lower()
    guessed_content_type = str(mimetypes.guess_type(str(file_name or ""))[0] or "").strip().lower()
    generic_types = {
        "",
        "application/octet-stream",
        "binary/octet-stream",
        "application/download",
        "application/x-download",
    }
    previewable_guess_types = {
        "application/pdf",
        "image/png",
        "image/jpeg",
        "image/jpg",
        "image/webp",
        "image/gif",
        "image/bmp",
        "image/svg+xml",
    }

    if raw_content_type in generic_types:
        return guessed_content_type or "application/octet-stream"
    if raw_content_type in {"text/plain", "application/plain"} and guessed_content_type in previewable_guess_types:
        return guessed_content_type
    return raw_content_type or guessed_content_type or "application/octet-stream"


def _sniff_document_content_type(file_bytes: bytes) -> str:
    payload = bytes(file_bytes or b"")
    if not payload:
        return ""

    header = payload[:4096]
    lowered_header = header.lower()
    if b"%PDF" in header[:1024]:
        return "application/pdf"
    if payload.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if payload.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if payload.startswith((b"GIF87a", b"GIF89a")):
        return "image/gif"
    if payload.startswith(b"RIFF") and payload[8:12] == b"WEBP":
        return "image/webp"
    if payload.startswith(b"BM"):
        return "image/bmp"
    if b"<svg" in lowered_header:
        return "image/svg+xml"
    return ""


def _resolve_inline_document_content_type(file_name: str, content_type: Any, file_bytes: bytes) -> str:
    sniffed_content_type = _sniff_document_content_type(file_bytes)
    if sniffed_content_type:
        return sniffed_content_type
    return _normalize_inline_document_content_type(file_name, content_type)


def _build_inline_document_response(
    file_name: str,
    content: Any,
    content_type: Any,
) -> Response:
    normalized_name = _extract_document_filename(file_name) or Path(str(file_name or "document")).name or "document"
    file_bytes = bytes(content or b"")
    resolved_content_type = _resolve_inline_document_content_type(normalized_name, content_type, file_bytes)
    headers = {"Content-Disposition": f'inline; filename="{normalized_name}"'}
    return Response(content=file_bytes, media_type=resolved_content_type, headers=headers)


def _build_inline_document_response_from_remote_url(
    remote_url: str,
    *,
    timeout_seconds: int = 20,
) -> Response | None:
    normalized_url = str(remote_url or "").strip()
    if not normalized_url:
        return None

    try:
        parsed_url = urlparse(normalized_url)
    except Exception:
        return None
    if parsed_url.scheme not in {"http", "https"}:
        return None

    request_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "Accept": "*/*",
    }

    try:
        remote_request = UrlRequest(normalized_url, headers=request_headers)
        with urlopen(remote_request, timeout=max(5, int(timeout_seconds))) as remote_response:
            response_bytes = remote_response.read()
            response_headers = getattr(remote_response, "headers", None)
            response_content_type = str(
                (response_headers.get("Content-Type") if response_headers else "")
                or ""
            ).split(";", 1)[0].strip()
            response_disposition = str(
                (response_headers.get("Content-Disposition") if response_headers else "")
                or ""
            ).strip()
            final_url = str(getattr(remote_response, "geturl", lambda: normalized_url)() or normalized_url).strip()
    except (HTTPError, URLError, OSError, ValueError):
        return None

    if not response_bytes:
        return None

    file_name = _extract_filename_from_content_disposition(response_disposition)
    if not file_name:
        parsed_final_url = urlparse(final_url or normalized_url)
        file_name = _extract_document_filename(parsed_final_url.path)
    if not file_name:
        file_name = _extract_document_filename(normalized_url)
    if not file_name:
        file_name = "document"

    if not Path(file_name).suffix:
        guessed_ext = str(mimetypes.guess_extension(response_content_type or "") or "").strip()
        if guessed_ext:
            file_name = f"{file_name}{guessed_ext}"

    return _build_inline_document_response(
        file_name=file_name,
        content=response_bytes,
        content_type=response_content_type,
    )


def _extract_filename_from_content_disposition(raw_value: Any) -> str:
    text = str(raw_value or "").strip()
    if not text:
        return ""

    matched = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^\";]+)"?', text, flags=re.IGNORECASE)
    if not matched:
        return ""
    extracted = unquote(str(matched.group(1) or "").strip())
    return Path(extracted).name.strip()


def _extract_document_response_payload(document_response: Response) -> tuple[bytes, str, str]:
    if isinstance(document_response, FileResponse):
        response_path = Path(str(getattr(document_response, "path", "") or "")).resolve()
        if not response_path.exists() or not response_path.is_file():
            raise HTTPException(status_code=404, detail="File not found")
        content_bytes = response_path.read_bytes()
        content_type = str(document_response.media_type or mimetypes.guess_type(str(response_path))[0] or "").strip()
        file_name = response_path.name
        return content_bytes, content_type, file_name

    body_bytes = bytes(getattr(document_response, "body", b"") or b"")
    media_type = str(document_response.media_type or "").strip()
    if not media_type:
        media_type = str(document_response.headers.get("content-type") or "").split(";", 1)[0].strip()
    file_name = _extract_filename_from_content_disposition(document_response.headers.get("content-disposition"))
    return body_bytes, media_type, file_name


def _build_pdf_only_response_from_document_response(
    document_response: Response,
    fallback_file_name: str = "document.pdf",
) -> Response:
    content_bytes, content_type, file_name = _extract_document_response_payload(document_response)
    sniffed_content_type = _sniff_document_content_type(content_bytes)
    normalized_content_type = str(content_type or "").split(";", 1)[0].strip().lower()

    if sniffed_content_type != "application/pdf":
        if normalized_content_type != "application/pdf" or b"%PDF" not in content_bytes[:1024]:
            raise HTTPException(status_code=415, detail="Requested document is not a PDF")

    resolved_name = _extract_document_filename(file_name) or _extract_document_filename(fallback_file_name) or "document.pdf"
    if not resolved_name.lower().endswith(".pdf"):
        resolved_name = f"{Path(resolved_name).stem or 'document'}.pdf"

    headers = {"Content-Disposition": f'inline; filename="{resolved_name}"'}
    return Response(content=content_bytes, media_type="application/pdf", headers=headers)


def _decode_document_reference_text_payload(file_bytes: bytes) -> str:
    payload = bytes(file_bytes or b"")
    if not payload:
        return ""
    if len(payload) > 4096:
        return ""

    def _reference_candidates(raw_text: str) -> list[str]:
        candidates: list[str] = []
        seen: set[str] = set()

        def add_candidate(candidate_value: Any) -> None:
            text = str(candidate_value or "").strip()
            if not text:
                return
            for variant in (text, text.strip("\"'")):
                normalized_variant = str(variant or "").strip()
                if not normalized_variant:
                    continue
                marker = normalized_variant.lower()
                if marker in seen:
                    continue
                seen.add(marker)
                candidates.append(normalized_variant)

        add_candidate(raw_text)

        try:
            parsed = json.loads(raw_text)
        except (json.JSONDecodeError, TypeError, ValueError):
            parsed = None

        if isinstance(parsed, (str, int, float)):
            add_candidate(parsed)
        elif isinstance(parsed, list):
            for item in parsed:
                if isinstance(item, (str, int, float)):
                    add_candidate(item)

        if raw_text.startswith("[") and raw_text.endswith("]"):
            for part in raw_text[1:-1].split(","):
                add_candidate(part)

        return candidates

    for encoding in ("utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be"):
        try:
            text = payload.decode(encoding, errors="strict").strip()
        except UnicodeDecodeError:
            continue
        if not text:
            continue

        for candidate in _reference_candidates(text):
            lowered = candidate.lower()
            if lowered.startswith(("http://", "https://", "/mock-files/")):
                return candidate
            if "/" in candidate or "\\" in candidate:
                return candidate
            if len(candidate) <= 256 and all(ch.isalnum() or ch in "-_. " for ch in candidate):
                return candidate

    return ""


def _looks_like_reference_text_payload(file_bytes: bytes) -> bool:
    return bool(_decode_document_reference_text_payload(file_bytes))


def _has_valid_document_signature(file_name: str, file_bytes: bytes) -> bool:
    payload = bytes(file_bytes or b"")
    if not payload:
        return False

    lowered_name = str(file_name or "").strip().lower()
    header = payload[:1024]
    if b"%PDF" in header:
        return True
    if lowered_name.endswith(".pdf"):
        return b"%PDF" in header
    if lowered_name.endswith(".png"):
        return payload.startswith(b"\x89PNG\r\n\x1a\n")
    if lowered_name.endswith((".jpg", ".jpeg")):
        return payload.startswith(b"\xff\xd8\xff")
    if lowered_name.endswith(".gif"):
        return payload.startswith((b"GIF87a", b"GIF89a"))
    if lowered_name.endswith(".webp"):
        return payload.startswith(b"RIFF") and payload[8:12] == b"WEBP"
    if _looks_like_reference_text_payload(payload):
        return False
    return True


def _cache_uploaded_document(storage_key: str, file_name: str, content: bytes, content_type: str) -> None:
    payload = {
        "content": content or b"",
        "content_type": content_type or "application/octet-stream",
    }

    normalized_key, normalized_name = _normalize_mock_file_reference(storage_key)
    if normalized_key:
        mock_files[normalized_key] = payload
    if normalized_name:
        mock_files[normalized_name] = payload

    legacy_name = Path(str(file_name or "")).name.strip()
    if legacy_name:
        mock_files[legacy_name] = payload


def _store_uploaded_document(
    record_id: str,
    field_name: str,
    file_name: str,
    content: bytes,
    content_type: str,
) -> str:
    normalized_name = Path(str(file_name or "")).name.strip()
    if not normalized_name:
        return ""

    storage_key = "/".join(
        (
            _sanitize_uploaded_file_segment(record_id, "record"),
            _sanitize_uploaded_file_segment(field_name, "document"),
            normalized_name,
        )
    )
    disk_path = _resolve_uploaded_file_disk_path(storage_key)
    if disk_path is not None:
        disk_path.parent.mkdir(parents=True, exist_ok=True)
        disk_path.write_bytes(content or b"")

    resolved_content_type = _normalize_inline_document_content_type(normalized_name, content_type)
    _cache_uploaded_document(storage_key, normalized_name, content or b"", resolved_content_type)
    return f"/mock-files/{quote(storage_key, safe='/')}"


async def _extract_form_body_and_uploads(
    form_data: Any,
) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    body: dict[str, Any] = {}
    uploads: dict[str, dict[str, Any]] = {}

    for key, value in form_data.multi_items():
        key_name = str(key or "").strip()
        if not key_name:
            continue

        if isinstance(value, (UploadFile, StarletteUploadFile)):
            raw_filename = str(value.filename or "").strip()
            normalized_filename = Path(raw_filename).name
            if not normalized_filename:
                continue

            file_bytes = await value.read()
            uploads[key_name] = {
                "filename": normalized_filename,
                "content": file_bytes or b"",
                "content_type": value.content_type or "application/octet-stream",
            }
        else:
            body[key_name] = value

    return body, uploads

VALIDATOR_LOCK_GROUPS = (
    {
        "flag_keys": ("gstValidate", "GST Validate"),
        "field_keys": (
            "gstRegistrationStatus",
            "GSTStatus",
            "gstNumber",
            "GSTIN",
            "tradeNameGst",
            "TradeName",
            "taxpayerTypeGst",
            "TaxpayerTypeGST",
            "gstDocument",
            "GSTAttachment",
            "GSTCertificateName",
        ),
    },
    {
        "flag_keys": ("cinValidate", "CIN Validate"),
        "field_keys": (
            "company",
            "CompanyStatus",
            "cinNumber",
            "CIN",
            "cinDocument",
            "CINCerti",
            "Incorporation_Filename",
        ),
    },
    {
        "flag_keys": ("panValidate", "PAN Validate"),
        "field_keys": (
            "panStatus",
            "PANStatus",
            "panNumber",
            "PANNo",
            "businessEntityType",
            "BusinessEntityType",
            "tradeNamePan",
            "TradeNamePAN",
            "panDocument",
            "PANAttachment",
            "PAN_Card_Filename",
        ),
    },
    {
        "flag_keys": ("msmeValidate", "MSME Validate"),
        "field_keys": (
            "msmeStatus",
            "MSMEStatus",
            "udyamNumber",
            "UdyamNumber",
            "msmeCategory",
            "MSMECategory",
            "msmeIndustry",
            "MSMEIndustry",
            "msmeDocument",
            "MSMECerti",
            "MSME_Certificate_Filename",
        ),
    },
    {
        "flag_keys": ("pfValidate", "PF Validate", "esiValidate", "ESI Validate"),
        "field_keys": (
            "pfStatus",
            "PFRegistrationStatus",
            "pfNumber",
            "PFNumber",
            "pfDocument",
            "PFCerti",
            "PF_Certificate_Filename",
            "esiStatus",
            "ESIStatus",
            "esiNumber",
            "ESIN",
            "esiDocument",
            "ESICerti",
            "ESI_Filename",
        ),
    },
    {
        "flag_keys": ("addressValidate", "Address Validate"),
        "field_keys": (
            "addressLane1",
            "Address1",
            "addressLane2",
            "Address2",
            "addressLane3",
            "Address3",
            "Adddress3",
            "country",
            "Country",
            "state",
            "State",
            "StateName",
            "district",
            "District",
            "pincode",
            "Pincode",
        ),
    },
    {
        "flag_keys": ("bankValidate", "Bank Validate"),
        "field_keys": (
            "bankPaymentMethod",
            "BankPaymentMethod",
            "ifscCode",
            "IFSC_Code",
            "bankName",
            "BankName",
            "branchName",
            "BranchName",
            "bankAccountNumber",
            "BankAccount",
            "cancelledCheque",
            "CancelledCheque",
            "Cancelled_Cheque_Filename",
        ),
    },
)

IMPORT_FORM_FIELD_TO_COLUMN = {
    "companyName": "companyName",
    "vendorName": "vendorName",
    "vendorCategory": "vendorCategory",
    "vendorType": "VendorType",
    "rubaminApproverHod": "RubaminApproverHod",
    "addressLane1": "addressLane1",
    "addressLane2": "addressLane2",
    "addressLane3": "addressLane3",
    "district": "district",
    "state": "state",
    "country": "country",
    "pincode": "pincode",
    "place": "Place",
    "currency": "Currency",
    "paymentTerms": "PaymentTerms",
    "incoTerms": "IncoTerms",
    "serviceProvideFor": "serviceProvideFor",
    "materialDealingsIn": "materialDealingsIn",
    "form10fStatus": "Form10FStatus",
    "bankPaymentMethod": "bankPaymentMethod",
    "beneficiaryName": "BeneficiaryName",
    "bankName": "bankName",
    "branchName": "branchName",
    "bankAccountNumber": "bankAccountNumber",
    "swiftCode": "SwiftCode",
    "iban": "IBAN",
    "contactPersonName": "contactPersonName",
    "contactPersonDesignation": "contactPersonDesignation",
    "contactPersonEmail": "contactPersonEmail",
    "contactPersonMobile": "contactPersonMobile",
    "alternativePersonName": "alternativePersonName",
    "alternativePersonDesignation": "alternativePersonDesignation",
    "alternativePersonEmail": "alternativePersonEmail",
    "alternativePersonMobile": "alternativePersonMobile",
    "alternativePersonDate": "alternativePersonDate",
    "alternativePersonPlace": "alternativePersonPlace",
    "agreeInformationAccuracy": "agreeInformationAccuracy",
}

IMPORT_FORM_UPLOAD_TO_COLUMN = {
    "salesContract": "SalesContract",
    "form10fDocument": "Form10FDocument",
    "trcDocument": "TRCDocument",
    "importDocument1": "ImportDocument1",
    "importDocument2": "ImportDocument2",
    "importDocument3": "ImportDocument3",
    "bankDocument": "ImportBankDocument",
}


def _document_source_aliases() -> dict[str, list[str]]:
    alias_map = {key: list(values) for key, values in DOCUMENT_FIELD_ALIASES.items()}
    for field_name, column_name in IMPORT_FORM_UPLOAD_TO_COLUMN.items():
        alias_map.setdefault(field_name, [])
        if column_name not in alias_map[field_name]:
            alias_map[field_name].append(column_name)
    return alias_map


def _get_document_candidate_keys(field_name: str) -> list[str]:
    normalized = str(field_name or "").strip()
    if not normalized:
        return []

    alias_map = _document_source_aliases()
    canonical = normalized
    lowered = normalized.lower()
    for candidate, aliases in alias_map.items():
        candidate_names = [candidate, *aliases]
        if any(str(name or "").strip().lower() == lowered for name in candidate_names):
            canonical = candidate
            break

    keys: list[str] = []
    seen_keys: set[str] = set()

    def add_key(candidate: Any) -> None:
        key_name = str(candidate or "").strip()
        if not key_name:
            return
        marker = key_name.lower()
        if marker in seen_keys:
            return
        seen_keys.add(marker)
        keys.append(key_name)

    for key in [canonical, *alias_map.get(canonical, []), normalized]:
        add_key(key)

        raw_key = str(key or "").strip()
        if not raw_key:
            continue

        base_key = raw_key
        lowered_key = raw_key.lower()
        if lowered_key.endswith("_filename"):
            base_key = raw_key[: -len("_FileName")]
        elif lowered_key.endswith("filename"):
            base_key = raw_key[: -len("FileName")]
        base_key = base_key.rstrip("_")
        if not base_key:
            continue

        add_key(base_key)
        add_key(f"{base_key}_FileName")
        add_key(f"{base_key}FileName")
        add_key(f"{base_key}_ContentType")
        add_key(f"{base_key}ContentType")

    return keys


def _extract_document_filename(raw_value: Any) -> str:
    value = str(raw_value or "").strip()
    if not value:
        return ""
    if value.startswith("/mock-files/"):
        value = value.split("/mock-files/", 1)[1]
    value = value.split("?", 1)[0]
    _, normalized_name = _normalize_mock_file_reference(value)
    return normalized_name

IMPORT_REQUIRED_FIELDS = (
    "companyName",
    "vendorName",
    "vendorCategory",
    "vendorType",
    "rubaminApproverHod",
    "addressLane1",
    "addressLane2",
    "country",
    "state",
    "district",
    "pincode",
    "currency",
    "paymentTerms",
    "incoTerms",
    "serviceProvideFor",
    "materialDealingsIn",
    "form10fStatus",
    "contactPersonName",
    "contactPersonEmail",
    "contactPersonMobile",
)

PAN_REGEX = re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$")
GST_REGEX = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")
CIN_REGEX = re.compile(r"^.{1,50}$")
PF_REGEX = re.compile(r"^[A-Z0-9][A-Z0-9/-]{5,29}$")
ESI_REGEX = re.compile(r"^.{1,50}$")
UDYAM_REGEX = re.compile(r"^UDYAM-[A-Z0-9]{2}-[0-9]{2}-[0-9]{7}$")
EMAIL_ADDRESS_REGEX = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def _extract_email_candidates(value: Any) -> list[str]:
    raw = str(value or "").strip()
    if not raw:
        return []

    candidates: list[str] = []
    seen: set[str] = set()

    matches = EMAIL_ADDRESS_REGEX.findall(raw)
    if matches:
        for item in matches:
            normalized = str(item or "").strip().lower()
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            candidates.append(normalized)
        return candidates

    for token in re.split(r"[;,\\s]+", raw):
        cleaned = str(token or "").strip().strip("<>()[]{}'\"").lower()
        if "@" not in cleaned or "." not in cleaned:
            continue
        if cleaned in seen:
            continue
        seen.add(cleaned)
        candidates.append(cleaned)

    return candidates

UIPATH_FIELDS: list[dict[str, str]] = [
    {"name": "Adddress3", "type": "string", "dataType": "System.String"},
    {"name": "Address1", "type": "string", "dataType": "System.String"},
    {"name": "Address2", "type": "string", "dataType": "System.String"},
    {"name": "Answer1", "type": "string", "dataType": "System.String"},
    {"name": "Answer2", "type": "string", "dataType": "System.String"},
    {"name": "Answer3", "type": "string", "dataType": "System.String"},
    {"name": "Answer4", "type": "string", "dataType": "System.String"},
    {"name": "Answer5", "type": "string", "dataType": "System.String"},
    {"name": "Answer6", "type": "string", "dataType": "System.String"},
    {"name": "Answer7", "type": "string", "dataType": "System.String"},
    {"name": "Answer8", "type": "string", "dataType": "System.String"},
    {"name": "Answer9", "type": "string", "dataType": "System.String"},
    {"name": "Approver_DT", "type": "object", "dataType": "entity.VendorMaster"},
    {"name": "BankAccount", "type": "string", "dataType": "System.String"},
    {"name": "BankName", "type": "string", "dataType": "System.String"},
    {"name": "BankPaymentMethod", "type": "string", "dataType": "System.String"},
    {"name": "BranchName", "type": "string", "dataType": "System.String"},
    {"name": "BusinessEntity", "type": "string", "dataType": "System.String"},
    {"name": "BusinessEntityType", "type": "string", "dataType": "System.String"},
    {"name": "Cancelled_Cheque_Filename", "type": "string", "dataType": "System.String"},
    {"name": "Cancelled_Cheque_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "CancelledCheque", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "CIN", "type": "string", "dataType": "System.String"},
    {"name": "CINCerti", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "CodeofConduct", "type": "boolean", "dataType": "System.Boolean"},
    {"name": "CompanyName", "type": "string", "dataType": "System.String"},
    {"name": "CompanyStatus", "type": "string", "dataType": "System.String"},
    {"name": "ContactEmailID", "type": "string", "dataType": "System.String"},
    {"name": "ContactPerson", "type": "string", "dataType": "System.String"},
    {"name": "CorporateLocation", "type": "string", "dataType": "System.String"},
    {"name": "Customer_List_Filename", "type": "string", "dataType": "System.String"},
    {"name": "Customer_List_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "CustomerDt", "type": "object", "dataType": "entity.CustomerMaster"},
    {"name": "CustomerMAster_Entity", "type": "object", "dataType": "entity.CustomerMaster"},
    {"name": "CustomerRec", "type": "object", "dataType": "entity.CustomerMaster"},
    {"name": "District", "type": "string", "dataType": "System.String"},
    {"name": "EcoVadis", "type": "string", "dataType": "System.String"},
    {"name": "ECOVadis_Filename", "type": "string", "dataType": "System.String"},
    {"name": "ECOVadis_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "EntityType", "type": "object", "dataType": "System.Object"},
    {"name": "ESI_Filename", "type": "string", "dataType": "System.String"},
    {"name": "ESI_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "ESICerti", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "ESIN", "type": "string", "dataType": "System.String"},
    {"name": "ESIStatus", "type": "string", "dataType": "System.String"},
    {"name": "GSTAttachment", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "GSTCertificateFile_Test", "type": "object", "dataType": "System.Object"},
    {"name": "GSTCertificateName", "type": "string", "dataType": "System.String"},
    {"name": "GSTCertificateName_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "GSTIN", "type": "string", "dataType": "System.String"},
    {"name": "GSTStatus", "type": "string", "dataType": "System.String"},
    {"name": "IFSC_Code", "type": "string", "dataType": "System.String"},
    {"name": "Incoporationdate", "type": "string", "dataType": "System.String"},
    {"name": "Incorporation_Filename", "type": "string", "dataType": "System.String"},
    {"name": "Incorporation_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "InhouseMachine", "type": "string", "dataType": "System.String"},
    {"name": "InhouseMachineAttchment", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "ISO_Certificate", "type": "list", "dataType": "System.Collections.Generic.List"},
    {"name": "ISO_File1_Name", "type": "string", "dataType": "System.String"},
    {"name": "ISO_File1_Name_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "ISO_File2_Name", "type": "string", "dataType": "System.String"},
    {"name": "ISO_File2_Name_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "ISO_File3_Name", "type": "string", "dataType": "System.String"},
    {"name": "ISO_File3_Name_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "ISO_File4_Name", "type": "string", "dataType": "System.String"},
    {"name": "ISO_File4_Name_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "Machineries_List_Filename", "type": "string", "dataType": "System.String"},
    {"name": "Machineries_List_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "MaterialDealing", "type": "string", "dataType": "System.String"},
    {"name": "MSME_Certificate_Filename", "type": "string", "dataType": "System.String"},
    {"name": "MSME_Certificate_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "MSMECategory", "type": "string", "dataType": "System.String"},
    {"name": "MSMECerti", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "MSMEIndustry", "type": "string", "dataType": "System.String"},
    {"name": "MSMEStatus", "type": "string", "dataType": "System.String"},
    {"name": "NumberofPlant", "type": "string", "dataType": "System.String"},
    {"name": "OHSAS_Code", "type": "string", "dataType": "System.String"},
    {"name": "OHSAS_Filename", "type": "string", "dataType": "System.String"},
    {"name": "OHSAS_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "Other_Filename", "type": "string", "dataType": "System.String"},
    {"name": "Other_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "OtherCerti", "type": "string", "dataType": "System.String"},
    {"name": "PAN_Card_Filename", "type": "string", "dataType": "System.String"},
    {"name": "PAN_Card_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "PANAttachment", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "PANNo", "type": "string", "dataType": "System.String"},
    {"name": "PANStatus", "type": "string", "dataType": "System.String"},
    {"name": "PF_Certificate_Filename", "type": "string", "dataType": "System.String"},
    {"name": "PF_Certificate_Filename_Downloaded", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "PFCerti", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "PFNumber", "type": "string", "dataType": "System.String"},
    {"name": "PFRegistrationStatus", "type": "string", "dataType": "System.String"},
    {"name": "Pincode", "type": "string", "dataType": "System.String"},
    {"name": "Previous_Year", "type": "string", "dataType": "System.String"},
    {"name": "PY_Turnover", "type": "string", "dataType": "System.String"},
    {"name": "RecordID", "type": "guid", "dataType": "System.Guid"},
    {"name": "Sample", "type": "object", "dataType": "System.Object"},
    {"name": "servicedealing", "type": "string", "dataType": "System.String"},
    {"name": "State", "type": "string", "dataType": "System.String"},
    {"name": "StateName", "type": "string", "dataType": "System.String"},
    {"name": "SupplierCode", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "TopCustomer", "type": "string", "dataType": "System.String"},
    {"name": "TopCustomerAttachment", "type": "object", "dataType": "Apps.Controls.AppsFile"},
    {"name": "TotalEmployees", "type": "string", "dataType": "System.String"},
    {"name": "TradeName", "type": "string", "dataType": "System.String"},
    {"name": "TradeNamePAN", "type": "string", "dataType": "System.String"},
    {"name": "UdyamDistrict", "type": "string", "dataType": "System.String"},
    {"name": "UdyamLast7Digit", "type": "string", "dataType": "System.String"},
    {"name": "UdyamNumber", "type": "string", "dataType": "System.String"},
    {"name": "UdyamState", "type": "string", "dataType": "System.String"},
    {"name": "Vendor_Category", "type": "string", "dataType": "System.String"},
    {"name": "VendoreMaster_Entity", "type": "object", "dataType": "entity.VendorMaster"},
    {"name": "Vendorname", "type": "string", "dataType": "System.String"},
    {"name": "VendorRec", "type": "object", "dataType": "entity.VendorMaster"},
]


def now_iso() -> str:
    """Return current time in Indian Standard Time (IST) as ISO string."""
    ist = timezone(timedelta(hours=5, minutes=30))
    return datetime.now(ist).isoformat()


def to_int(value: Any) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def empty_to_none(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    return value


def validate_business_identifier_fields(body: dict[str, Any]) -> None:
    def normalize(key: str, uppercase: bool = True) -> str | None:
        raw = empty_to_none(body.get(key))
        if raw is None:
            return None
        normalized = str(raw).strip()
        if uppercase:
            normalized = normalized.upper()
        body[key] = normalized
        return normalized

    validations = [
        ("panNumber", PAN_REGEX, "PAN", True),
        ("gstNumber", GST_REGEX, "GST", True),
        ("cinNumber", CIN_REGEX, "CIN", True),
        ("pfNumber", PF_REGEX, "PF", True),
        ("esiNumber", ESI_REGEX, "ESI", False),
        ("udyamNumber", UDYAM_REGEX, "Udyam Number", True),
    ]

    for key, pattern, label, uppercase in validations:
        value = normalize(key, uppercase=uppercase)
        if value is None:
            continue
        if not pattern.fullmatch(value):
            raise HTTPException(status_code=400, detail=f"Invalid {label} format")


def to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.lower() == "true"
    return bool(value)


def truthy_flag(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


def get_validator_locked_update_keys(record: dict[str, Any]) -> set[str]:
    if not record:
        return set()

    locked_keys: set[str] = set()
    for group in VALIDATOR_LOCK_GROUPS:
        flag_keys = group.get("flag_keys", ())
        if not any(key in record and truthy_flag(record.get(key)) for key in flag_keys):
            continue
        for key in group.get("field_keys", ()):
            key_name = str(key or "").strip()
            if key_name:
                locked_keys.add(key_name.lower())
    return locked_keys


def hash_password(password: str) -> str:
    digest = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return f"sha256${digest}"


def check_password(password: str, hashed: str) -> bool:
    return hash_password(password) == hashed


def make_token(user: dict[str, Any], password_change_required: bool = False) -> str:
    payload = {
        "id": user["id"],
        "email": user["email"],
        "role": user["role"],
        "botName": normalize_bot_name(user.get("botName") or "Vendor Bot"),
        "passwordChangeRequired": bool(password_change_required),
        "exp": datetime.now(timezone.utc) + timedelta(days=7),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)


def get_user_min(user_id: Any) -> dict[str, Any] | None:
    uid = to_int(user_id)
    if uid is None:
        return None
    user = users.get(uid)
    if not user:
        return None
    return {"id": user["id"], "name": user["name"], "email": user["email"]}


def hydrate_form(record: dict[str, Any]) -> dict[str, Any]:
    form = dict(record)
    form["vendor"] = get_user_min(form.get("vendorId"))
    form["buyer"] = get_user_min(form.get("buyerId"))
    form["hod"] = get_user_min(form.get("hodId"))
    return form


def sorted_forms(filters: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    filters = filters or {}
    rows: list[dict[str, Any]] = []
    for form in vendor_forms.values():
        matched = all(form.get(k) == v for k, v in filters.items())
        if matched:
            rows.append(hydrate_form(form))
    rows.sort(key=lambda x: x.get("createdAt", ""), reverse=True)
    return rows


def _normalize_record_id(value: Any) -> str:
    return str(value or "").strip()


def _normalize_record_key(value: Any) -> str:
    return _normalize_record_id(value).lower()


def _record_keys_from_form(form: dict[str, Any]) -> list[str]:
    return [
        _normalize_record_key(form.get("recordId")),
        _normalize_record_key(form.get("id")),
    ]


def _run_gst_extraction_inline(gst_number: str) -> dict[str, Any] | None:
    """
    Run the GST Extraction Process Python code directly (no HTTP API).
    Returns the first result dict on success, or None on failure.
    """
    gst_number = str(gst_number or "").strip()
    if not gst_number:
        return None

    try:
        base_dir = Path(__file__).resolve().parent.parent
        gst_dir = (
            base_dir
            / "Data Validation"
            / "GST Extraction Process"
        )
        gst_main_path = gst_dir / "main.py"
        if not gst_main_path.is_file():
            print(f"[GST] main.py not found at {gst_main_path}")
            return None

        # Ensure the GST folder is on sys.path so its own imports
        # like `from ocr_module import ...` work exactly as when
        # running `python main.py` in that directory.
        import sys

        added_path = False
        if str(gst_dir) not in sys.path:
            sys.path.insert(0, str(gst_dir))
            added_path = True

        try:
            spec = importlib.util.spec_from_file_location(
                "gst_extraction_main", str(gst_main_path)
            )
            if spec is None or spec.loader is None:
                print("[GST] Failed to create import spec for main.py")
                return None
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)  # type: ignore[arg-type]

            if not hasattr(module, "main"):
                print("[GST] main.py does not expose a main() function")
                return None

            results = module.main([gst_number], save_to_excel=False)  # type: ignore[attr-defined]
            if isinstance(results, list) and results:
                first = results[0]
                if isinstance(first, dict):
                    return first
            return None
        finally:
            if added_path and str(gst_dir) in sys.path:
                sys.path.remove(str(gst_dir))
    except Exception as exc:
        print(f"[GST] Inline extraction failed: {exc}")
        return None


def _normalize_financial_year_value(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip())
    if not text:
        return ""
    match = re.search(r"(20\d{2})\D*(\d{2,4})", text)
    if not match:
        return text
    start_year = int(match.group(1))
    end_raw = str(match.group(2))
    end_year = (int(end_raw) % 100) if len(end_raw) == 4 else int(end_raw)
    return f"{start_year}-{end_year:02d}"


def _run_gst_filing_table_extraction_inline(
    gst_number: str,
    financial_year: str | None = None,
) -> dict[str, Any] | None:
    """
    Run the GST filing-table extraction Python code directly (no HTTP API).
    Returns the first result dict on success, or None on failure.
    """
    gst_number = str(gst_number or "").strip()
    if not gst_number:
        return None

    try:
        base_dir = Path(__file__).resolve().parent.parent
        gst_dir = (
            base_dir
            / "Data Validation"
            / "GST Extraction Process"
        )
        extractor_path = gst_dir / "filing_table_extractor.py"
        if not extractor_path.is_file():
            print(f"[GST-FILING] filing_table_extractor.py not found at {extractor_path}")
            return None

        import sys

        added_path = False
        if str(gst_dir) not in sys.path:
            sys.path.insert(0, str(gst_dir))
            added_path = True

        try:
            spec = importlib.util.spec_from_file_location(
                "gst_filing_table_extractor", str(extractor_path)
            )
            if spec is None or spec.loader is None:
                print("[GST-FILING] Failed to create import spec for filing_table_extractor.py")
                return None
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)  # type: ignore[arg-type]

            if not hasattr(module, "main"):
                print("[GST-FILING] filing_table_extractor.py does not expose a main() function")
                return None

            normalized_financial_year = _normalize_financial_year_value(financial_year)
            try:
                if normalized_financial_year:
                    results = module.main(  # type: ignore[attr-defined]
                        [gst_number],
                        save_to_excel=False,
                        financial_year=normalized_financial_year,
                    )
                else:
                    results = module.main([gst_number], save_to_excel=False)  # type: ignore[attr-defined]
            except TypeError:
                results = module.main([gst_number], save_to_excel=False)  # type: ignore[attr-defined]
            if isinstance(results, list) and results:
                first = results[0]
                if isinstance(first, dict):
                    return first
            return None
        finally:
            if added_path and str(gst_dir) in sys.path:
                sys.path.remove(str(gst_dir))
    except Exception as exc:
        print(f"[GST-FILING] Inline extraction failed: {exc}")
        return None


def _normalize_inline_status_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _is_registered_gst_status(value: Any) -> bool:
    return _normalize_inline_status_text(value) in {
        "registered",
        "yes",
        "y",
        "true",
        "1",
        "active",
        "valid",
    }


def _extract_gst_taxpayer_type(gst_payload: dict[str, Any] | None) -> str:
    if not isinstance(gst_payload, dict):
        return ""

    for key in (
        "Taxpayer Type",
        "TaxpayerTypeGST",
        "taxpayerTypeGst",
        "TaxpayerType",
        "taxpayerType",
    ):
        value = gst_payload.get(key)
        text = str(value or "").strip()
        if text:
            return text
    return ""


def _run_msme_extraction_inline(udyam_number: str) -> dict[str, Any] | None:
    """
    Run the MSME Udyam extraction Python code directly (no HTTP API).
    Returns a result dict on success, or None on failure.
    """
    udyam_number = str(udyam_number or "").strip()
    if not udyam_number:
        return None

    try:
        base_dir = Path(__file__).resolve().parent.parent
        msme_dir = (
            base_dir
            / "Data Validation"
            / "MSME Validation"
        )
        msme_path = msme_dir / "Udyam Data Extract.py"
        if not msme_path.is_file():
            print(f"[MSME] Udyam Data Extract.py not found at {msme_path}")
            return None

        import sys

        added_path = False
        if str(msme_dir) not in sys.path:
            sys.path.insert(0, str(msme_dir))
            added_path = True

        try:
            # Import module as a package-style name (replace spaces and extension)
            module_name = "udyam_data_extract"
            spec = importlib.util.spec_from_file_location(module_name, str(msme_path))
            if spec is None or spec.loader is None:
                print("[MSME] Failed to create import spec for Udyam Data Extract.py")
                return None
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)  # type: ignore[arg-type]

            if not hasattr(module, "run_single_udyam"):
                print("[MSME] Udyam Data Extract.py does not expose run_single_udyam()")
                return None

            return module.run_single_udyam(udyam_number)  # type: ignore[attr-defined]
        finally:
            if added_path and str(msme_dir) in sys.path:
                sys.path.remove(str(msme_dir))
    except Exception as exc:
        print(f"[MSME] Inline extraction failed: {exc}")
        return None


GST_VALIDATION_CHANGE_FIELDS: tuple[dict[str, Any], ...] = (
    {
        "label": "Trade Name as per GST",
        "keys": ("tradeNameGst", "TradeName", "TradeNameGST"),
    },
    {
        "label": "Taxpayer Type",
        "keys": ("taxpayerTypeGst", "TaxpayerTypeGST", "Taxpayer Type"),
    },
    {
        "label": "GST Valid",
        "keys": ("GST Valid", "GSTValid", "gstValid"),
    },
)

MSME_VALIDATION_CHANGE_FIELDS: tuple[dict[str, Any], ...] = (
    {
        "label": "MSME Category",
        "keys": ("msmeCategory", "MSMECategory"),
    },
    {
        "label": "MSME Industry",
        "keys": ("msmeIndustry", "MSMEIndustry"),
    },
    {
        "label": "Udyam Registration Date",
        "keys": ("udyamRegistrationDate", "UdyamRegistrationDate", "MSMEDate", "msmeDate"),
    },
)

GST_VALIDATION_CHANGE_NOTE_KEYS = ("gstValidationChanges", "GSTValidationChanges")
MSME_VALIDATION_CHANGE_NOTE_KEYS = ("msmeValidationChanges", "MSMEValidationChanges")


def _validation_note_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (bytes, bytearray, memoryview)):
        return ""
    if isinstance(value, (list, dict, tuple, set)):
        try:
            return json.dumps(value, ensure_ascii=True)
        except Exception:
            return str(value).strip()
    return str(value).strip()


def _first_update_value(updates: dict[str, Any], *keys: str) -> tuple[Any, bool]:
    if not isinstance(updates, dict):
        return None, False
    lowered = {str(key or "").strip().lower(): key for key in updates.keys()}
    for key in keys:
        actual_key = lowered.get(str(key or "").strip().lower())
        if actual_key is not None:
            return updates.get(actual_key), True
    return None, False


def _build_validation_change_note(
    previous_local_record: dict[str, Any] | None,
    previous_db_record: dict[str, Any] | None,
    updates: dict[str, Any],
    fields: tuple[dict[str, Any], ...],
    section: str,
) -> str:
    rows: list[dict[str, str]] = []
    for field in fields:
        keys = tuple(str(key or "").strip() for key in field.get("keys", ()) if str(key or "").strip())
        if not keys:
            continue
        new_value, has_new_value = _first_update_value(updates, *keys)
        if not has_new_value:
            continue

        previous_text = _validation_note_text(
            _first_existing_record_value(previous_local_record, previous_db_record, *keys)
        )
        current_text = _validation_note_text(new_value)
        if previous_text == current_text:
            continue

        rows.append(
            {
                "field": str(field.get("label") or keys[0]),
                "previousValue": previous_text,
                "newValue": current_text,
            }
        )

    if not rows:
        return ""

    return json.dumps(
        {
            "section": section,
            "generatedAt": now_iso(),
            "rows": rows,
        },
        ensure_ascii=True,
    )


def _run_vendor_post_submit_enrichment(record_id: str) -> None:
    normalized_record_id = _normalize_record_id(record_id)
    if not normalized_record_id:
        return

    local_record = find_local_vendor_record_by_record_id(normalized_record_id)
    db_record = get_vendor_record_by_record_id(normalized_record_id)
    if not local_record and not db_record:
        print(f"[ENRICHMENT] Vendor record not found for recordId={normalized_record_id}")
        return

    approver_status = str(
        _first_existing_record_value(
            local_record,
            db_record,
            "approverStatus",
            "ApproverStatus",
            "status",
            "Status",
        )
        or ""
    ).strip().lower()

    should_enrich = approver_status in {
        "pending for user approval",
        "pendign for user approval",
        "pending for validator approval",
        "pendign for validator approval",
    }
    if not should_enrich:
        return

    updates: dict[str, Any] = {}

    gst_status_raw = str(
        _first_existing_record_value(local_record, db_record, "gstRegistrationStatus", "GSTStatus") or ""
    ).strip().lower()
    gst_number_raw = str(
        _first_existing_record_value(local_record, db_record, "gstNumber", "GSTIN") or ""
    ).strip().upper()
    should_run_gst_api = (
        gst_number_raw
        and (
            "registered" in gst_status_raw
            or gst_status_raw in {"yes", "y", "true", "1"}
        )
    )
    if should_run_gst_api:
        print(f"[GST] Running post-submit extraction for GSTIN={gst_number_raw!r}")
        gst_data = _run_gst_extraction_inline(gst_number_raw) or {}
        trade_name = str(gst_data.get("Trade Name") or "").strip()
        taxpayer_type = _extract_gst_taxpayer_type(gst_data)
        status_value = str(gst_data.get("GSTIN / UIN  Status") or "").strip().lower()
        if status_value in {"valid", "active"}:
            if trade_name:
                updates["tradeNameGst"] = trade_name
                updates["TradeName"] = trade_name
            if taxpayer_type:
                updates["taxpayerTypeGst"] = taxpayer_type
                updates["TaxpayerTypeGST"] = taxpayer_type

    msme_status_raw = str(
        _first_existing_record_value(local_record, db_record, "MSMEStatus", "msmeStatus") or ""
    ).strip().lower()
    udyam_number_raw = str(
        _first_existing_record_value(local_record, db_record, "UdyamNumber", "udyamNumber") or ""
    ).strip().upper()
    should_run_msme = (
        udyam_number_raw
        and (
            "registered" in msme_status_raw
            or msme_status_raw in {"yes", "y", "true", "1"}
        )
    )
    if should_run_msme:
        print(f"[MSME] Running post-submit extraction for Udyam={udyam_number_raw!r}")
        msme_data = _run_msme_extraction_inline(udyam_number_raw) or {}
        msme_category = str(msme_data.get("MSME Category") or "").strip()
        msme_industry = str(msme_data.get("MSME Industry") or "").strip()
        udyam_reg_date = str(msme_data.get("Date of Udyam Registration") or "").strip()
        if msme_category:
            updates["msmeCategory"] = msme_category
            updates["MSMECategory"] = msme_category
        if msme_industry:
            updates["msmeIndustry"] = msme_industry
            updates["MSMEIndustry"] = msme_industry
        if udyam_reg_date:
            updates["udyamRegistrationDate"] = udyam_reg_date
            updates["UdyamRegistrationDate"] = udyam_reg_date

    gst_change_note = _build_validation_change_note(
        local_record,
        db_record,
        updates,
        GST_VALIDATION_CHANGE_FIELDS,
        "GST Details",
    )
    if gst_change_note:
        for note_key in GST_VALIDATION_CHANGE_NOTE_KEYS:
            updates[note_key] = gst_change_note

    msme_change_note = _build_validation_change_note(
        local_record,
        db_record,
        updates,
        MSME_VALIDATION_CHANGE_FIELDS,
        "MSME Details",
    )
    if msme_change_note:
        for note_key in MSME_VALIDATION_CHANGE_NOTE_KEYS:
            updates[note_key] = msme_change_note

    if updates:
        update_local_vendor_record_by_record_id(normalized_record_id, updates)

        try:
            columns_to_ensure = [
                key
                for key in (
                    "TaxpayerTypeGST",
                    *GST_VALIDATION_CHANGE_NOTE_KEYS,
                    *MSME_VALIDATION_CHANGE_NOTE_KEYS,
                )
                if str(updates.get(key) or "").strip()
            ]
            if columns_to_ensure:
                ensure_vendor_columns(columns_to_ensure)
            update_vendor_record_by_record_id(normalized_record_id, updates)
        except Exception as exc:
            print(f"[ENRICHMENT] Failed to persist enrichment updates for {normalized_record_id}: {exc}")

    try:
        _store_vendor_code_of_conduct_pdf(normalized_record_id)
    except Exception as exc:
        print(f"[ENRICHMENT] Failed to generate Code of Conduct PDF for {normalized_record_id}: {exc}")


def find_local_vendor_record_by_record_id(record_id: str) -> dict[str, Any] | None:
    record_key = _normalize_record_key(record_id)
    if not record_key:
        return None

    for invite in invite_forms.values():
        if _normalize_record_key(invite.get("recordId")) == record_key:
            return dict(invite)

    for form in vendor_forms.values():
        if record_key in _record_keys_from_form(form):
            return dict(form)

    return None


def update_local_vendor_status_by_record_id(record_id: str, approver_status: str) -> bool:
    record_key = _normalize_record_key(record_id)
    if not record_key:
        return False

    updated = False
    updated_at = now_iso()

    for invite in invite_forms.values():
        if _normalize_record_key(invite.get("recordId")) != record_key:
            continue
        invite["approverStatus"] = approver_status
        invite["status"] = approver_status
        invite["updatedAt"] = updated_at
        updated = True

    for form in vendor_forms.values():
        if record_key not in _record_keys_from_form(form):
            continue
        form["approverStatus"] = approver_status
        form["updatedAt"] = updated_at
        updated = True

    return updated


def update_local_vendor_record_by_record_id(record_id: str, updates: dict[str, Any]) -> bool:
    record_key = _normalize_record_key(record_id)
    if not record_key or not isinstance(updates, dict) or not updates:
        return False

    protected_keys = {"recordid", "record", "id"}
    updated = False
    updated_at = now_iso()

    for invite in invite_forms.values():
        if _normalize_record_key(invite.get("recordId")) != record_key:
            continue
        for key, value in updates.items():
            key_name = str(key or "").strip()
            if not key_name:
                continue
            if key_name.lower() in protected_keys:
                continue
            invite[key_name] = value
        invite["updatedAt"] = updated_at
        updated = True

    for form in vendor_forms.values():
        if record_key not in _record_keys_from_form(form):
            continue
        for key, value in updates.items():
            key_name = str(key or "").strip()
            if not key_name:
                continue
            if key_name.lower() in protected_keys:
                continue
            form[key_name] = value
        form["updatedAt"] = updated_at
        updated = True

    return updated


def _extract_record_email(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""
    email_keys = (
        "contactPersonEmail",
        "ContactPersonEmail",
        "ContactEmailID",
        "contactEmailId",
        "vendorEmail",
        "VendorEmail",
        "invitedByEmail",
    )
    for key in email_keys:
        candidates = _extract_email_candidates(record.get(key))
        if candidates:
            return candidates[0]
    return ""


def _extract_record_assignee_email(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""

    email_keys = (
        "rubaminContactPersonEmail",
        "Rubamin Contact Person Email",
        "rubaminContactPerson",
        "Rubamin Contact Person",
        "invitedByEmail",
        "InvitedByEmail",
        "Invited By Email",
        "buyerEmail",
        "BuyerEmail",
        "contactPersonEmail",
        "ContactPersonEmail",
        "ContactEmailID",
        "contactEmailId",
    )
    for key in email_keys:
        candidates = _extract_email_candidates(record.get(key))
        if candidates:
            return candidates[0]

    buyer_id = record.get("buyerId")
    _, buyer_email = _lookup_user_name_email_by_id(buyer_id)
    buyer_candidates = _extract_email_candidates(buyer_email)
    return buyer_candidates[0] if buyer_candidates else ""


def _record_is_assigned_to_email(record: dict[str, Any] | None, target_email: str) -> bool:
    if not isinstance(record, dict):
        return False

    normalized_target = str(target_email or "").strip().lower()
    if not normalized_target:
        return False

    email_keys = (
        "rubaminContactPersonEmail",
        "Rubamin Contact Person Email",
        "rubaminContactPerson",
        "Rubamin Contact Person",
        "invitedByEmail",
        "InvitedByEmail",
        "Invited By Email",
        "buyerEmail",
        "BuyerEmail",
        "contactPersonEmail",
        "ContactPersonEmail",
        "ContactEmailID",
        "contactEmailId",
    )
    for key in email_keys:
        for candidate in _extract_email_candidates(record.get(key)):
            if candidate == normalized_target:
                return True

    buyer_id = record.get("buyerId")
    _, buyer_email = _lookup_user_name_email_by_id(buyer_id)
    for candidate in _extract_email_candidates(buyer_email):
        if candidate == normalized_target:
            return True

    return False


def _extract_record_hod_email(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""
    hod_keys = (
        "rubaminApproverHod",
        "RubaminApproverHod",
        "Rubamin Approver HOD",
        "RubaminApprovalHod",
        "Rubamin Approval HOD",
        "RubaminApprovalHOD",
        "hodEmail",
        "HodEmail",
        "HODEmail",
        "HOD Email",
    )
    for key in hod_keys:
        for candidate in _extract_email_candidates(record.get(key)):
            if candidate:
                return candidate
        value = str(record.get(key) or "").strip().lower()
        if value:
            return value

    # Fallback for legacy/custom column spellings.
    for key, raw_value in record.items():
        normalized_key = re.sub(r"[^a-z0-9]+", "", str(key or "").strip().lower())
        if not normalized_key or "hod" not in normalized_key:
            continue
        if not any(token in normalized_key for token in ("email", "approver", "approval")):
            continue
        for candidate in _extract_email_candidates(raw_value):
            if candidate:
                return candidate
        value = str(raw_value or "").strip().lower()
        if value:
            return value
    return ""


def _extract_record_vendor_email(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""
    vendor_email_keys = (
        "vendorEmail",
        "VendorEmail",
        "contactPersonEmail",
        "ContactPersonEmail",
        "ContactEmailID",
        "contactEmailId",
    )
    for key in vendor_email_keys:
        value = str(record.get(key) or "").strip().lower()
        if value:
            return value
    return ""


def _extract_record_vendor_name(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""
    vendor_name_keys = (
        "vendorName",
        "Vendorname",
        "VendorName",
        "name",
        "Name",
    )
    for key in vendor_name_keys:
        value = str(record.get(key) or "").strip()
        if value:
            return value
    return ""


def _extract_record_buyer_email(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""

    buyer_email_keys = (
        "buyerEmail",
        "BuyerEmail",
        "rubaminContactPersonEmail",
        "Rubamin Contact Person Email",
        "invitedByEmail",
    )
    for key in buyer_email_keys:
        candidates = _extract_email_candidates(record.get(key))
        if candidates:
            return candidates[0]

    buyer_id = record.get("buyerId")
    _, buyer_email = _lookup_user_name_email_by_id(buyer_id)
    buyer_candidates = _extract_email_candidates(buyer_email)
    return buyer_candidates[0] if buyer_candidates else ""


def _extract_record_buyer_name(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""

    buyer_name_keys = (
        "rubaminContactPerson",
        "Rubamin Contact Person",
    )
    for key in buyer_name_keys:
        value = str(record.get(key) or "").strip()
        if value:
            return value

    buyer_id = record.get("buyerId")
    buyer_name, _ = _lookup_user_name_email_by_id(buyer_id)
    return str(buyer_name or "").strip()


def _resolve_vendor_code_notification_details(
    local_record: dict[str, Any] | None,
    db_record: dict[str, Any] | None,
) -> tuple[str, str, str]:
    recipient_email = str(
        _first_existing_record_value(
            local_record,
            db_record,
            "rubaminContactPersonEmail",
            "Rubamin Contact Person Email",
            "buyerEmail",
            "BuyerEmail",
            "invitedByEmail",
            "InvitedByEmail",
            "Invited By Email",
        )
        or _extract_record_buyer_email(local_record)
        or _extract_record_buyer_email(db_record)
        or _extract_record_assignee_email(local_record)
        or _extract_record_assignee_email(db_record)
        or _extract_record_vendor_email(local_record)
        or _extract_record_vendor_email(db_record)
        or ""
    ).strip().lower()
    recipient_name = str(
        _first_existing_record_value(
            local_record,
            db_record,
            "rubaminContactPerson",
            "Rubamin Contact Person",
            "contactPersonName",
            "ContactPerson",
        )
        or _extract_record_buyer_name(local_record)
        or _extract_record_buyer_name(db_record)
        or _derive_name_from_email(recipient_email)
        or "User"
    ).strip()
    vendor_name = str(
        _first_existing_record_value(
            local_record,
            db_record,
            "vendorName",
            "Vendorname",
            "VendorName",
            "name",
            "Name",
        )
        or "Vendor"
    ).strip()
    return recipient_name, recipient_email, vendor_name


def _collect_record_access_emails(*records: dict[str, Any] | None) -> set[str]:
    allowed_emails: set[str] = set()
    email_keys = (
        "contactPersonEmail",
        "ContactPersonEmail",
        "ContactEmailID",
        "contactEmailId",
        "vendorEmail",
        "VendorEmail",
        "rubaminContactPersonEmail",
        "Rubamin Contact Person Email",
        "rubaminContactPerson",
        "Rubamin Contact Person",
        "invitedByEmail",
        "InvitedByEmail",
        "Invited By Email",
        "buyerEmail",
        "BuyerEmail",
    )

    for record in records:
        if not isinstance(record, dict):
            continue

        for key in email_keys:
            for candidate in _extract_email_candidates(record.get(key)):
                allowed_emails.add(candidate)

        buyer_id = record.get("buyerId")
        _, buyer_email = _lookup_user_name_email_by_id(buyer_id)
        for candidate in _extract_email_candidates(buyer_email):
            allowed_emails.add(candidate)

    return allowed_emails


def _authorize_vendor_record_access(
    local_record: dict[str, Any] | None,
    db_record: dict[str, Any] | None,
    user: dict[str, Any] | None,
) -> None:
    if user is None:
        return

    user_email = str(user.get("email") or "").strip().lower()

    if user_has_any_role(user, "admin", "validator"):
        return

    candidate_record_ids: set[str] = set()
    for record in (local_record, db_record):
        if not isinstance(record, dict):
            continue
        for key in ("recordId", "RecordID", "record", "Record", "id", "Id", "ID"):
            value = _normalize_record_id(record.get(key))
            if value:
                candidate_record_ids.add(value.lower())

    if user_has_any_role(user, "hod"):
        allowed_hod_emails = {
            email
            for email in (
                _extract_record_hod_email(local_record),
                _extract_record_hod_email(db_record),
            )
            if email
        }
        if user_email and allowed_hod_emails and user_email in allowed_hod_emails:
            return
        if not user_has_any_role(user, "user", "buyer"):
            if user_email and candidate_record_ids:
                try:
                    for row in get_vendor_history_records(user):
                        visible_record = _normalize_record_id(row.get("record")).lower()
                        if visible_record and visible_record in candidate_record_ids:
                            return
                except Exception:
                    pass
            raise HTTPException(status_code=403, detail="Access denied for this record")

    allowed_emails = _collect_record_access_emails(local_record, db_record)
    if user_email and user_email in allowed_emails:
        return

    if user_email and candidate_record_ids:
        if candidate_record_ids:
            for row in get_vendor_records_by_assignee_email(user_email):
                db_row_record = _normalize_record_id(row.get("record")).lower()
                if db_row_record and db_row_record in candidate_record_ids:
                    return
        try:
            for row in get_vendor_history_records(user):
                visible_record = _normalize_record_id(row.get("record")).lower()
                if visible_record and visible_record in candidate_record_ids:
                    return
        except Exception:
            pass

    raise HTTPException(status_code=403, detail="Access denied for this record")


def _extract_prescreen_answers(record: dict[str, Any] | None) -> dict[str, str]:
    if not isinstance(record, dict):
        return {}
    answers: dict[str, str] = {}
    for idx in range(1, 10):
        candidates = (
            f"q{idx}",
            f"Q{idx}",
            f"answer{idx}",
            f"Answer{idx}",
        )
        value = ""
        for key in candidates:
            raw = record.get(key)
            if raw is None:
                continue
            value = str(raw).strip()
            if value:
                break
        if value:
            answers[f"q{idx}"] = value
    return answers


def _lookup_user_name_email_by_id(user_id: Any) -> tuple[str, str]:
    uid = to_int(user_id)
    if uid is None:
        return "", ""

    local_user = users.get(uid)
    if local_user:
        return (
            str(local_user.get("name") or "").strip(),
            str(local_user.get("email") or "").strip(),
        )

    for db_user in get_db_users():
        if to_int(db_user.get("id")) == uid:
            return (
                str(db_user.get("name") or "").strip(),
                str(db_user.get("email") or "").strip(),
            )

    return "", ""


def _format_person_name(name_value: Any) -> str:
    raw_name = str(name_value or "").strip()
    if not raw_name:
        return ""

    if "@" in raw_name:
        raw_name = raw_name.split("@", 1)[0].strip()

    normalized = re.sub(r"[._]+", " ", raw_name)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    if not normalized:
        return ""

    words: list[str] = []
    for raw_word in normalized.split(" "):
        word = raw_word.strip()
        if not word:
            continue

        parts = [segment for segment in re.split(r"([\-'])", word) if segment != ""]
        formatted_parts: list[str] = []
        for part in parts:
            if part in {"-", "'"}:
                formatted_parts.append(part)
                continue
            if part.isupper() and len(part) <= 3:
                formatted_parts.append(part)
            else:
                formatted_parts.append(part[:1].upper() + part[1:].lower())

        words.append("".join(formatted_parts))

    return " ".join(words).strip()


def _derive_name_from_email(email_value: Any) -> str:
    email_text = str(email_value or "").strip()
    if "@" in email_text:
        email_text = email_text.split("@", 1)[0].strip()
    return _format_person_name(email_text)


def resolve_rubamin_contact_person_details(body: dict[str, Any], user: dict[str, Any]) -> dict[str, Any]:
    default_email = str(user.get("email") or "").strip()
    default_name = str(user.get("name") or "").strip() or _derive_name_from_email(default_email)
    contact_name = default_name
    contact_email = default_email
    normalized_role = normalize_role(user.get("role"))
    if normalized_role == "vendor":
        buyer_name, buyer_email = _lookup_user_name_email_by_id(body.get("buyerId"))
        if buyer_name:
            contact_name = buyer_name
        elif buyer_email and not contact_name:
            contact_name = _derive_name_from_email(buyer_email)
        if buyer_email:
            contact_email = buyer_email

    if not contact_name and contact_email:
        contact_name = _derive_name_from_email(contact_email)

    department_value = empty_to_none(body.get("department"))
    if isinstance(department_value, str):
        department_value = department_value.strip() or None

    return {
        "name": contact_name or None,
        "email": contact_email or None,
        "department": department_value,
    }


def _merge_rubamin_contact_with_existing_record(
    contact: dict[str, Any],
    existing_record: dict[str, Any] | None,
) -> dict[str, Any]:
    merged = dict(contact)
    if not isinstance(existing_record, dict):
        return merged

    if not merged.get("name"):
        merged["name"] = empty_to_none(
            existing_record.get("rubaminContactPerson")
            or existing_record.get("Rubamin Contact Person")
        )
    if not merged.get("email"):
        merged["email"] = empty_to_none(
            existing_record.get("rubaminContactPersonEmail")
            or existing_record.get("Rubamin Contact Person Email")
        )
    if not merged.get("department"):
        merged["department"] = empty_to_none(
            existing_record.get("rubaminContactPersonDepartment")
            or existing_record.get("Rubamin Contact Person Department")
        )
    return merged


def _first_existing_record_value(
    existing_local_record: dict[str, Any] | None,
    existing_db_record: dict[str, Any] | None,
    *keys: str,
) -> Any:
    for record in (existing_local_record, existing_db_record):
        if not isinstance(record, dict):
            continue
        for key in keys:
            value = record.get(key)
            if value is None:
                continue
            if isinstance(value, str):
                trimmed = value.strip()
                if trimmed:
                    return trimmed
            else:
                return value
    return None


def _get_existing_vendor_record_or_raise(record_id: str) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    record_value = _normalize_record_id(record_id)
    if not record_value:
        raise HTTPException(status_code=400, detail="recordId is required")

    local_record = find_local_vendor_record_by_record_id(record_value)
    db_record = get_vendor_record_by_record_id(record_value)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Invalid or expired recordId")
    return local_record, db_record


def _persist_vendor_record_by_record_id(
    record_id: str,
    payload: dict[str, Any],
    *,
    existing_db_record: dict[str, Any] | None,
) -> str:
    if update_vendor_record_by_record_id(record_id, payload):
        return record_id
    if existing_db_record:
        raise RuntimeError("Vendor record found in database but update failed")

    inserted_record_id = insert_vendor_master_data(payload) or record_id
    # When there is an existing DB record (invited/vendor flow), we must
    # preserve the invited recordId. For brand‑new internal records, allow
    # SQL to generate its own RecordID without raising.
    if existing_db_record and _normalize_record_id(inserted_record_id) != _normalize_record_id(record_id):
        raise RuntimeError("Vendor record insert did not preserve the invited recordId")
    return inserted_record_id


def _build_public_registration_url(request: Request, registration_path: str) -> str:
    path = str(registration_path or "").strip()
    if not path:
        return ""
    if not path.startswith("/"):
        path = f"/{path}"
    base = PUBLIC_BASE_URL or str(request.base_url or "").strip().rstrip("/")
    if not base:
        return path
    return f"{base}{path}"


def _resolve_resubmission_target_from_status(current_status: Any) -> tuple[str, str] | tuple[None, None]:
    normalized_status = str(current_status or "").strip().lower()
    if "reject" not in normalized_status:
        return None, None

    if "hod" in normalized_status:
        return "hod", "Pending For HOD Approval"
    if "validator" in normalized_status or "validat" in normalized_status:
        return "validator", "Pending For Validator Approval"
    if "user" in normalized_status or "buyer" in normalized_status:
        return "user", "Pending For User Approval"
    return "user", "Pending For User Approval"


def _send_stage_approval_notifications(
    *,
    stage: str,
    record_id: str,
    vendor_name: str,
    submitted_on: str,
    approver_status: str,
    dashboard_url: str,
    remarks: str | None = None,
    merged_record: dict[str, Any] | None = None,
    local_record: dict[str, Any] | None = None,
    db_record: dict[str, Any] | None = None,
    preferred_contact_name: str | None = None,
    preferred_contact_email: str | None = None,
    entity_label: str = "Vendor",
) -> tuple[int, str | None]:
    sent_count = 0
    errors: list[str] = []
    remarks_text = str(remarks or "").strip()

    normalized_stage = str(stage or "").strip().lower()
    if normalized_stage == "validator":
        validator_contacts = _get_validator_contacts()
        if not validator_contacts:
            errors.append("Validator email not configured")
        for contact in validator_contacts:
            contact_email = str(contact.get("email") or "").strip()
            if not contact_email:
                continue
            contact_name = str(contact.get("name") or "").strip() or _derive_name_from_email(contact_email)
            try:
                send_vendor_submission_approval_email(
                    contact_person_name=contact_name,
                    contact_person_email=contact_email,
                    record_id=record_id,
                    vendor_name=vendor_name,
                    submitted_on=submitted_on,
                    approver_status=approver_status,
                    dashboard_url=dashboard_url,
                    remarks=remarks_text,
                    entity_label=entity_label,
                )
                sent_count += 1
            except Exception as exc:
                errors.append(f"{contact_email}: {exc}")
    elif normalized_stage == "hod":
        hod_email = (
            _extract_record_hod_email(merged_record)
            or _extract_record_hod_email(local_record)
            or _extract_record_hod_email(db_record)
        )
        if not hod_email:
            errors.append("HOD email not configured on this record")
        else:
            try:
                send_vendor_submission_approval_email(
                    contact_person_name=_derive_name_from_email(hod_email),
                    contact_person_email=hod_email,
                    record_id=record_id,
                    vendor_name=vendor_name,
                    submitted_on=submitted_on,
                    approver_status=approver_status,
                    dashboard_url=dashboard_url,
                    remarks=remarks_text,
                    entity_label=entity_label,
                )
                sent_count += 1
            except Exception as exc:
                errors.append(str(exc))
    else:
        contact_email = (
            str(preferred_contact_email or "").strip()
            or _extract_record_assignee_email(merged_record)
            or _extract_record_assignee_email(local_record)
            or _extract_record_assignee_email(db_record)
        )
        contact_name = (
            str(preferred_contact_name or "").strip()
            or _extract_record_buyer_name(merged_record)
            or _extract_record_buyer_name(local_record)
            or _extract_record_buyer_name(db_record)
            or _derive_name_from_email(contact_email)
        )
        if not contact_email:
            errors.append("User approver email not configured on this record")
        else:
            try:
                send_vendor_submission_approval_email(
                    contact_person_name=contact_name,
                    contact_person_email=contact_email,
                    record_id=record_id,
                    vendor_name=vendor_name,
                    submitted_on=submitted_on,
                    approver_status=approver_status,
                    dashboard_url=dashboard_url,
                    remarks=remarks_text,
                    entity_label=entity_label,
                )
                sent_count += 1
            except Exception as exc:
                errors.append(str(exc))

    return sent_count, "; ".join(errors) if errors else None


def _build_vendor_invite_email_text(
    vendor_name: str,
    record_id: str,
    registration_url: str,
    entity_label: str = "Vendor",
) -> str:
    entity_title = str(entity_label or "").strip() or "Vendor"
    return (
        f'Dear {vendor_name},\n\n'
        f"We request you to kindly fill out the {entity_title} Creation Form using the link below to initiate your onboarding "
        f'process using the Record ID: <b>{record_id}</b>\n\n'
        f"Click here to access {entity_title} Creation Form: {registration_url}\n\n"
        "Please ensure all mandatory fields are completed and supporting documents are uploaded as required. "
        "This will help us process your registration smoothly and without delay.\n\n"
        "For any queries or assistance, feel free to reach out to us.\n\n"
        "Thank you for your cooperation.\n\n"
        "Best Regards,\n"
        "Sharp & Tannan"
    )


def _build_vendor_invite_email_html(
    vendor_name: str,
    record_id: str,
    registration_url: str,
    entity_label: str = "Vendor",
) -> str:
    safe_vendor_name = html_escape(vendor_name)
    safe_record_id = html_escape(record_id)
    safe_registration_url = html_escape(registration_url, quote=True)
    entity_title = html_escape(str(entity_label or "").strip() or "Vendor")
    return (
        '<html><body style="font-family: Verdana, Geneva, sans-serif;">'
        f"<p>Dear {safe_vendor_name},</p>"
        f"<p>We request you to kindly fill out the {entity_title} Creation Form using the link below to initiate your "
        f'onboarding process using the Record ID: <b>{safe_record_id}</b></p>'
        f'<p><a href="{safe_registration_url}">Click here to access {entity_title} Creation Form</a></p>'
        "<p>Please ensure all mandatory fields are completed and supporting documents are uploaded as required. "
        "This will help us process your registration smoothly and without delay.</p>"
        "<p>For any queries or assistance, feel free to reach out to us.</p>"
        "<p>Thank you for your cooperation.</p>"
        "<p>Best Regards,<br>Sharp &amp; Tannan</p>"
        "</body></html>"
    )


def _send_smtp_message(message: EmailMessage, recipients: list[str]) -> None:
    smtp_user = str(SMTP_USERNAME or "").strip()
    sender_email = str(SMTP_SENDER_EMAIL or "").strip()
    if not smtp_user:
        smtp_user = sender_email
    if not smtp_user:
        raise RuntimeError("SMTP_USERNAME/SMTP_SENDER_EMAIL is not configured")
    if not SMTP_PASSWORD:
        raise RuntimeError("SMTP_PASSWORD is not configured")

    def auth_error_message(exc: smtplib.SMTPAuthenticationError, attempted_user: str) -> str:
        code = getattr(exc, "smtp_code", "")
        raw_error = getattr(exc, "smtp_error", b"")
        if isinstance(raw_error, (bytes, bytearray)):
            detail = raw_error.decode("utf-8", errors="replace").strip()
        else:
            detail = str(raw_error or "").strip()
        message_text = (
            f"SMTP authentication failed for '{attempted_user}'. "
            "Set a valid Gmail/SMTP app password for SMTP_USERNAME/SMTP_SENDER_EMAIL."
        )
        if code or detail:
            message_text = f"{message_text} ({code}) {detail}".strip()
        return message_text

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT_SECONDS) as smtp:
        smtp.ehlo()
        if SMTP_USE_STARTTLS:
            smtp.starttls()
            smtp.ehlo()
        try:
            smtp.login(smtp_user, SMTP_PASSWORD)
        except smtplib.SMTPAuthenticationError as exc:
            fallback_user = sender_email
            if fallback_user and fallback_user.lower() != smtp_user.lower():
                try:
                    smtp.login(fallback_user, SMTP_PASSWORD)
                except smtplib.SMTPAuthenticationError as fallback_exc:
                    raise RuntimeError(auth_error_message(fallback_exc, fallback_user)) from fallback_exc
            else:
                raise RuntimeError(auth_error_message(exc, smtp_user)) from exc
        smtp.send_message(message, from_addr=SMTP_SENDER_EMAIL, to_addrs=recipients)


def send_vendor_invite_email(
    *,
    vendor_name: str,
    vendor_email: str,
    record_id: str,
    registration_url: str,
    cc_email: str | None = None,
    entity_label: str = "Vendor",
) -> None:
    if not SMTP_PASSWORD:
        raise RuntimeError("SMTP_PASSWORD is not configured")

    to_email = str(vendor_email or "").strip()
    entity_title = str(entity_label or "").strip() or "Vendor"
    if not to_email:
        raise RuntimeError(f"{entity_title} email is required for sending invite email")

    cc_value = str(cc_email or "").strip()
    recipients = [to_email]
    if cc_value and cc_value.lower() != to_email.lower():
        recipients.append(cc_value)

    display_vendor_name = _format_person_name(vendor_name) or _derive_name_from_email(to_email) or entity_title

    message = EmailMessage()
    message["Subject"] = f"{entity_title} Creation Form | Record ID: {record_id}"
    message["From"] = SMTP_SENDER_EMAIL
    message["To"] = to_email
    if cc_value:
        message["Cc"] = cc_value
    message.set_content(
        _build_vendor_invite_email_text(
            display_vendor_name,
            record_id,
            registration_url,
            entity_label=entity_title,
        )
    )
    message.add_alternative(
        _build_vendor_invite_email_html(
            display_vendor_name,
            record_id,
            registration_url,
            entity_label=entity_title,
        ),
        subtype="html",
    )

    _send_smtp_message(message, recipients)


def _format_email_datetime(value: Any) -> str:
    raw_value = str(value or "").strip()
    if not raw_value:
        return ""
    try:
        parsed = datetime.fromisoformat(raw_value.replace("Z", "+00:00"))
    except ValueError:
        return raw_value
    ist = timezone(timedelta(hours=5, minutes=30))
    if parsed.tzinfo is None:
        # App/SQL timestamps without timezone are persisted in IST.
        parsed = parsed.replace(tzinfo=ist)
    return parsed.astimezone(ist).strftime("%Y-%m-%d %H:%M IST")


def _build_vendor_submission_approval_email_text(
    contact_person_name: str,
    record_id: str,
    vendor_name: str,
    submitted_on: str,
    approver_status: str,
    dashboard_url: str,
    remarks: str = "",
    entity_label: str = "Vendor",
) -> str:
    entity_title = str(entity_label or "").strip() or "Vendor"
    remarks_text = str(remarks or "").strip() or "N/A"
    return (
        f"Dear {contact_person_name},\n\n"
        f"New {entity_title} Request has been submitted for approval.\n\n"
        "Request Details:\n\n"
        f"Request ID : <b>{record_id}</b>\n"
        f"{entity_title} Name : <b>{vendor_name} </b>\n"
        f"Submitted On : <b>{submitted_on}</b>\n"
        f"Approver Status : <b>{approver_status}</b>\n"
        f"Remarks : <b>{remarks_text}</b>\n\n"
        "We request you to kindly review and approve the request by visiting the link below:\n"
        f"Click here to approve the request: {dashboard_url}\n\n"
        "Please feel free to reach out in case of any queries.\n\n"
        "Thank you for your support.\n\n"
        "Best regards,"
    )


def _build_vendor_submission_approval_email_html(
    contact_person_name: str,
    record_id: str,
    vendor_name: str,
    submitted_on: str,
    approver_status: str,
    dashboard_url: str,
    remarks: str = "",
    entity_label: str = "Vendor",
) -> str:
    entity_title = html_escape(str(entity_label or "").strip() or "Vendor")
    remarks_text = html_escape(str(remarks or "").strip() or "N/A")
    return (
        '<html><body style="font-family: Verdana, Geneva, sans-serif;">'
        f"<p>Dear {html_escape(contact_person_name)},</p>"
        f"<p>New {entity_title} Request has been submitted for approval.</p>"
        "<p>Request Details:</p>"
        "<table style=\"border-collapse:collapse\">"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Request ID :</strong></td><td><b>{html_escape(record_id)}</b></td></tr>"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>{entity_title} Name :</strong></td><td><b>{html_escape(vendor_name)}</b></td></tr>"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Submitted On :</strong></td><td><b>{html_escape(submitted_on)}</b></td></tr>"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Approver Status :</strong></td><td><b>{html_escape(approver_status)}</b></td></tr>"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Remarks :</strong></td><td><b>{remarks_text}</b></td></tr>"
        "</table>"
        "<p>We request you to kindly review and approve the request by visiting the link below:</p>"
        f'<p><a href="{html_escape(dashboard_url, quote=True)}">Click here to approve the request</a></p>'
        "<p>Please feel free to reach out in case of any queries.</p>"
        "<p>Thank you for your support.</p>"
        "<p>Best regards,</p>"
        "</body></html>"
    )


def send_vendor_submission_approval_email(
    *,
    contact_person_name: str,
    contact_person_email: str,
    record_id: str,
    vendor_name: str,
    submitted_on: str,
    approver_status: str,
    dashboard_url: str,
    remarks: str = "",
    entity_label: str = "Vendor",
) -> None:
    if not SMTP_PASSWORD:
        raise RuntimeError("SMTP_PASSWORD is not configured")

    to_email = str(contact_person_email or "").strip()
    if not to_email:
        raise RuntimeError("Rubamin contact person email is required for approval email")

    display_name = _format_person_name(contact_person_name) or _derive_name_from_email(to_email) or "Team"
    submitted_on_text = _format_email_datetime(submitted_on) or str(submitted_on or "").strip()
    status_text = str(approver_status or "").strip()
    remarks_text = str(remarks or "").strip()
    entity_title = str(entity_label or "").strip() or "Vendor"
    recipients = [to_email]

    message = EmailMessage()
    message["Subject"] = f"{entity_title} Form Submitted for your Approval"
    message["From"] = SMTP_SENDER_EMAIL
    message["To"] = to_email
    message.set_content(
        _build_vendor_submission_approval_email_text(
            display_name,
            record_id,
            vendor_name,
            submitted_on_text,
            status_text,
            dashboard_url,
            remarks_text,
            entity_label=entity_title,
        )
    )
    message.add_alternative(
        _build_vendor_submission_approval_email_html(
            display_name,
            record_id,
            vendor_name,
            submitted_on_text,
            status_text,
            dashboard_url,
            remarks_text,
            entity_label=entity_title,
        ),
        subtype="html",
    )

    _send_smtp_message(message, recipients)


def _build_vendor_code_created_email_text(
    recipient_name: str,
    vendor_name: str,
    vendor_code: str,
) -> str:
    return (
        f"Dear {recipient_name},\n\n"
        "Vendor Code is created in the system for your request. The details of the vendor are as follows:\n\n"
        f"Vendor Name: {vendor_name}\n"
        f"Vendor Code (SAP): {vendor_code}\n\n"
        "Best regards,\n"
        "Sharp & Tannan Associates"
    )


def _build_vendor_code_created_email_html(
    recipient_name: str,
    vendor_name: str,
    vendor_code: str,
) -> str:
    return (
        '<html><body style="font-family: Verdana, Geneva, sans-serif;">'
        f"<p>Dear {html_escape(recipient_name)},</p>"
        "<p>Vendor Code is created in the system for your request. The details of the vendor are as follows:</p>"
        "<table style=\"border-collapse:collapse\">"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Vendor Name:</strong></td><td>{html_escape(vendor_name)}</td></tr>"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Vendor Code (SAP):</strong></td><td>{html_escape(vendor_code)}</td></tr>"
        "</table>"
        "<p>Best regards,<br>Sharp &amp; Tannan Associates</p>"
        "</body></html>"
    )


def send_vendor_code_created_email(
    *,
    recipient_name: str,
    recipient_email: str,
    vendor_name: str,
    vendor_code: str,
    record_id: str,
) -> None:
    if not SMTP_PASSWORD:
        raise RuntimeError("SMTP_PASSWORD is not configured")

    to_email = str(recipient_email or "").strip().lower()
    if not to_email:
        raise RuntimeError("Recipient email is required for vendor code created email")

    display_name = _format_person_name(recipient_name) or _derive_name_from_email(to_email) or "User"
    display_vendor_name = str(vendor_name or "").strip() or "Vendor"
    display_vendor_code = str(vendor_code or "").strip()
    if not display_vendor_code:
        raise RuntimeError("Vendor code is required for vendor code created email")

    message = EmailMessage()
    message["Subject"] = f"Vendor Code Created in SAP | Record ID: {record_id}"
    message["From"] = SMTP_SENDER_EMAIL
    message["To"] = to_email
    message.set_content(
        _build_vendor_code_created_email_text(
            display_name,
            display_vendor_name,
            display_vendor_code,
        )
    )
    message.add_alternative(
        _build_vendor_code_created_email_html(
            display_name,
            display_vendor_name,
            display_vendor_code,
        ),
        subtype="html",
    )

    _send_smtp_message(message, [to_email])


def _build_vendor_rejection_email_text(
    recipient_name: str,
    remarks: str,
    record_id: str,
    update_form_url: str,
    rejected_by: str,
    entity_label: str = "Vendor",
) -> str:
    rejected_by_label = str(rejected_by or "").strip() or "review team"
    entity_title = str(entity_label or "").strip() or "Vendor"
    return (
        f"Dear {recipient_name},\n\n"
        f"Your {entity_title} Registration request has been rejected by the {rejected_by_label} with the remarks -\n\n"
        f"{remarks}\n\n"
        f"So we request you to kindly modify the form based on the provided remarks for the Record ID: {record_id}\n\n"
        f"Click here to access Update {entity_title} Form: {update_form_url}\n\n"
        "Please ensure all mandatory fields are completed as required. This will help us process your registration smoothly and without delay.\n\n"
        "For any queries or assistance, feel free to reach out to us.\n\n"
        "Thank you for your cooperation.\n\n"
        "Best Regards,\n"
        "Sharp & Tannan"
    )


def _build_vendor_rejection_email_html(
    recipient_name: str,
    remarks: str,
    record_id: str,
    update_form_url: str,
    rejected_by: str,
    entity_label: str = "Vendor",
) -> str:
    rejected_by_label = html_escape(str(rejected_by or "").strip() or "review team")
    entity_title = html_escape(str(entity_label or "").strip() or "Vendor")
    return (
        '<html><body style="font-family: Verdana, Geneva, sans-serif;">'
        f"<p>Dear {html_escape(recipient_name)},</p>"
        f"<p>Your {entity_title} Registration request has been rejected by the {rejected_by_label} with the remarks -</p>"
        f"<p><b>{html_escape(remarks)}</b></p>"
        f"<p>So we request you to kindly modify the form based on the provided remarks for the Record ID: <b>{html_escape(record_id)}</b></p>"
        f'<p><a href="{html_escape(update_form_url, quote=True)}">Click here to access Update {entity_title} Form</a></p>'
        "<p>Please ensure all mandatory fields are completed as required. This will help us process your registration smoothly and without delay.</p>"
        "<p>For any queries or assistance, feel free to reach out to us.</p>"
        "<p>Thank you for your cooperation.</p>"
        "<p>Best Regards,<br>Sharp &amp; Tannan</p>"
        "</body></html>"
    )


def send_vendor_rejection_email(
    *,
    recipient_name: str,
    recipient_email: str,
    record_id: str,
    remarks: str,
    update_form_url: str,
    rejected_by: str,
    cc_email: str | None = None,
    entity_label: str = "Vendor",
) -> None:
    if not SMTP_PASSWORD:
        raise RuntimeError("SMTP_PASSWORD is not configured")

    to_email = str(recipient_email or "").strip()
    if not to_email:
        raise RuntimeError("Recipient email is required for rejection email")

    cc_value = str(cc_email or "").strip()
    recipients = [to_email]
    if cc_value and cc_value.lower() != to_email.lower():
        recipients.append(cc_value)

    display_name = _format_person_name(recipient_name) or _derive_name_from_email(to_email) or "Team"
    remarks_text = str(remarks or "").strip() or "-"
    entity_title = str(entity_label or "").strip() or "Vendor"

    message = EmailMessage()
    message["Subject"] = f"{entity_title} Registration Request Rejected"
    message["From"] = SMTP_SENDER_EMAIL
    message["To"] = to_email
    if cc_value and cc_value.lower() != to_email.lower():
        message["Cc"] = cc_value
    message.set_content(
        _build_vendor_rejection_email_text(
            display_name,
            remarks_text,
            record_id,
            update_form_url,
            rejected_by,
            entity_label=entity_title,
        )
    )
    message.add_alternative(
        _build_vendor_rejection_email_html(
            display_name,
            remarks_text,
            record_id,
            update_form_url,
            rejected_by,
            entity_label=entity_title,
        ),
        subtype="html",
    )

    _send_smtp_message(message, recipients)


def _build_new_user_welcome_email_text(
    recipient_name: str,
    username: str,
    password: str,
    login_url: str,
) -> str:
    return (
        f"Dear {recipient_name},\n\n"
        "Welcome to Sharp & Tannan Vendor Portal.\n\n"
        "Your login credentials are:\n"
        f"Username: {username}\n"
        f"Password: {password}\n\n"
        f"Login URL: {login_url}\n\n"
        "For security, please change your password immediately after first login.\n\n"
        "Best Regards,\n"
        "Sharp & Tannan"
    )


def _build_new_user_welcome_email_html(
    recipient_name: str,
    username: str,
    password: str,
    login_url: str,
) -> str:
    return (
        '<html><body style="font-family: Verdana, Geneva, sans-serif;">'
        f"<p>Dear {html_escape(recipient_name)},</p>"
        "<p>Welcome to Sharp & Tannan Vendor Portal.</p>"
        "<p>Your login credentials are:</p>"
        "<table style=\"border-collapse:collapse\">"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Username:</strong></td><td>{html_escape(username)}</td></tr>"
        f"<tr><td style=\"padding:4px 12px 4px 0\"><strong>Password:</strong></td><td>{html_escape(password)}</td></tr>"
        "</table>"
        f'<p><strong>Login URL:</strong> <a href="{html_escape(login_url, quote=True)}">{html_escape(login_url)}</a></p>'
        "<p>For security, please change your password immediately after first login.</p>"
        "<p>Best Regards,<br>Sharp &amp; Tannan</p>"
        "</body></html>"
    )


def send_new_user_welcome_email(
    *,
    recipient_name: str,
    recipient_email: str,
    username: str,
    password: str,
    login_url: str,
) -> None:
    if not SMTP_PASSWORD:
        raise RuntimeError("SMTP_PASSWORD is not configured")

    to_email = str(recipient_email or "").strip()
    if not to_email:
        raise RuntimeError("Recipient email is required for welcome email")

    display_name = _format_person_name(recipient_name) or _derive_name_from_email(to_email) or "Team"
    login_link = str(login_url or "").strip() or "/web/login"

    message = EmailMessage()
    message["Subject"] = "Welcome to Sharp & Tannan Vendor Portal"
    message["From"] = SMTP_SENDER_EMAIL
    message["To"] = to_email
    message.set_content(
        _build_new_user_welcome_email_text(
            display_name,
            username,
            password,
            login_link,
        )
    )
    message.add_alternative(
        _build_new_user_welcome_email_html(
            display_name,
            username,
            password,
            login_link,
        ),
        subtype="html",
    )

    _send_smtp_message(message, [to_email])


def create_user(name: str, email: str, password: str, role: str, bot_name: str = "Vendor Bot") -> dict[str, Any]:
    uid = next(user_id_counter)
    user = {
        "id": uid,
        "name": name,
        "email": email,
        "password": hash_password(password),
        "role": role,
        "botName": normalize_bot_name(bot_name),
        "passwordChangeRequired": False,
        "createdAt": now_iso(),
    }
    users[uid] = user
    users_by_email[email.lower()] = uid
    return user


def ensure_local_user(
    name: str,
    email: str,
    role: str,
    preferred_id: Any = None,
    bot_name: str | None = None,
) -> dict[str, Any]:
    existing_id = users_by_email.get(email.lower())
    if existing_id:
        user = users[existing_id]
        user["name"] = name
        user["role"] = role
        if bot_name is not None:
            user["botName"] = normalize_bot_name(bot_name)
        return user

    preferred_int = to_int(preferred_id)
    if preferred_int is not None and preferred_int > 0 and preferred_int not in users:
        user = {
            "id": preferred_int,
            "name": name,
            "email": email,
            "password": "",
            "role": role,
            "botName": normalize_bot_name(bot_name or "Vendor Bot"),
            "passwordChangeRequired": False,
            "createdAt": now_iso(),
        }
        users[preferred_int] = user
        users_by_email[email.lower()] = preferred_int
        return user

    return create_user(name, email, "__sql_managed_user__", role, bot_name=bot_name or "Vendor Bot")


def user_from_token(token: str) -> dict[str, Any] | None:
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
    except jwt.PyJWTError:
        return None

    password_change_required = bool(payload.get("passwordChangeRequired"))
    token_bot_name = normalize_bot_name(payload.get("botName") or payload.get("botname") or "Vendor Bot")
    user_id = to_int(payload.get("id"))
    user = users.get(user_id or -1)
    if user:
        user["passwordChangeRequired"] = password_change_required
        if not user.get("botName"):
            user["botName"] = token_bot_name
        return user

    token_email = str(payload.get("email") or "").strip()
    token_role = str(payload.get("role") or "Vendor").strip() or "Vendor"
    if not token_email:
        return None

    token_name = token_email.split("@")[0] or "SQL User"
    resolved_user = ensure_local_user(
        name=token_name,
        email=token_email,
        role=token_role,
        preferred_id=user_id,
        bot_name=token_bot_name,
    )
    resolved_user["passwordChangeRequired"] = password_change_required
    return resolved_user


def _build_authenticated_user_payload(user: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": user["id"],
        "name": user.get("name"),
        "email": user["email"],
        "role": user["role"],
        "botName": normalize_bot_name(user.get("botName") or "Vendor Bot"),
        "passwordChangeRequired": bool(user.get("passwordChangeRequired")),
    }


def get_optional_current_user(
    credentials: HTTPAuthorizationCredentials | None = Depends(security),
) -> dict[str, Any] | None:
    if not credentials:
        return None
    if credentials.scheme.lower() != "bearer":
        raise HTTPException(status_code=401, detail="Unsupported authorization scheme")

    user = user_from_token(credentials.credentials)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid or expired token")

    return _build_authenticated_user_payload(user)


def get_current_user(
    credentials: HTTPAuthorizationCredentials | None = Depends(security),
) -> dict[str, Any]:
    user = get_optional_current_user(credentials)
    if not user:
        raise HTTPException(status_code=401, detail="No token provided")
    return user


def get_current_user_from_token_or_cookie(
    request: Request,
    credentials: HTTPAuthorizationCredentials | None = Depends(security),
) -> dict[str, Any]:
    if credentials:
        if credentials.scheme.lower() != "bearer":
            raise HTTPException(status_code=401, detail="Unsupported authorization scheme")
        user = user_from_token(credentials.credentials)
        if not user:
            raise HTTPException(status_code=401, detail="Invalid or expired token")
        return _build_authenticated_user_payload(user)

    cookie_user = get_cookie_session_user(request)
    return _build_authenticated_user_payload(cookie_user)


ROLE_ALIASES = {
    "admin": "admin",
    "administrator": "admin",
    "hod": "hod",
    "head of department": "hod",
    "user": "user",
    "customer user": "user",
    "vendor user": "user",
    "vendor": "vendor",
    "buyer": "buyer",
    "banking": "banking",
    "accounts": "accounts",
    "taxation": "taxation",
    "validator": "validator",
    "validator & user": "validator_user",
    "validator and user": "validator_user",
    "validator+user": "validator_user",
    "validator_user": "validator_user",
    "validator & hod": "validator_hod",
    "validator and hod": "validator_hod",
    "validator+hod": "validator_hod",
    "validator_hod": "validator_hod",
}


def normalize_role(value: Any) -> str:
    raw = str(value or "").strip().lower()
    return ROLE_ALIASES.get(raw, raw)


ROLE_CAPABILITY_MAP: dict[str, set[str]] = {
    "validator_user": {"validator", "user", "buyer"},
    "validator_hod": {"validator", "hod"},
}


def get_role_capabilities(value: Any) -> set[str]:
    normalized = normalize_role(value)
    if normalized == "admin":
        return {
            "admin",
            "validator",
            "hod",
            "user",
            "buyer",
            "vendor",
            "customer",
            "banking",
            "accounts",
            "taxation",
        }

    capabilities = set(ROLE_CAPABILITY_MAP.get(normalized, {normalized}))
    if "user" in capabilities:
        capabilities.add("buyer")
    if "buyer" in capabilities:
        capabilities.add("user")
    return capabilities


def user_has_any_role(user: dict[str, Any] | None, *roles: str) -> bool:
    if not isinstance(user, dict):
        return False
    capabilities = get_role_capabilities(user.get("role"))
    if "admin" in capabilities:
        return True
    allowed = {normalize_role(role) for role in roles}
    return any(role in capabilities for role in allowed)


def normalize_bot_name(value: Any) -> str:
    raw = str(value or "").strip().lower()
    if raw in {"both", "all", "vendor+customer", "vendor & customer"}:
        return "Both"
    if raw in {"customer bot", "customer", "cust", "cust bot"}:
        return "Customer Bot"
    # Default to vendor for unknown values to preserve backward compatibility.
    return "Vendor Bot"


def get_cookie_session_user(request: Request) -> dict[str, Any]:
    token = request.cookies.get(AUTH_COOKIE_NAME)
    if not token:
        raise HTTPException(status_code=401, detail="No active login session")
    user = user_from_token(token)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid or expired login session")
    return user


def require_role(*roles: str):
    def checker(user: dict[str, Any] = Depends(get_current_user)) -> dict[str, Any]:
        if not user_has_any_role(user, *roles):
            raise HTTPException(status_code=403, detail="Access denied. Insufficient permissions")
        return user

    return checker


def seed_users() -> None:
    return


@app.get("/")
def health() -> dict[str, str]:
    return {"message": "Server is running"}


@app.get("/api/db/health")
def db_health() -> dict[str, Any]:
    return get_db_status()


def _extract_postal_location(payload: Any) -> tuple[str, str, str] | None:
    if not isinstance(payload, list):
        return None

    for entry in payload:
        if not isinstance(entry, dict):
            continue

        status = str(entry.get("Status") or "").strip().lower()
        if status and status != "success":
            continue

        post_offices = entry.get("PostOffice")
        if not isinstance(post_offices, list):
            continue

        for office in post_offices:
            if not isinstance(office, dict):
                continue
            district = str(office.get("District") or "").strip()
            state = str(office.get("State") or "").strip()
            country = str(office.get("Country") or "").strip() or "India"
            if district and state:
                return district, state, country

    return None


def _extract_ifsc_location(payload: Any) -> tuple[str, str] | None:
    if not isinstance(payload, dict):
        return None

    bank_name = str(payload.get("BANK") or payload.get("Bank") or payload.get("bank") or "").strip()
    branch_name = str(payload.get("BRANCH") or payload.get("Branch") or payload.get("branch") or "").strip()
    if not bank_name or not branch_name:
        return None
    return bank_name, branch_name


@app.get("/api/pincode/{pincode}")
def lookup_pincode_details(pincode: str) -> dict[str, str]:
    normalized_pincode = re.sub(r"\D", "", str(pincode or ""))
    if len(normalized_pincode) != 6:
        raise HTTPException(status_code=400, detail="Pincode must be 6 digits")

    lookup_url = f"https://api.postalpincode.in/pincode/{normalized_pincode}"
    try:
        with urlopen(lookup_url, timeout=8) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except HTTPError as exc:
        status_code = exc.code if isinstance(exc.code, int) and exc.code > 0 else 502
        raise HTTPException(status_code=status_code, detail="Pincode service request failed") from exc
    except URLError as exc:
        raise HTTPException(status_code=502, detail="Pincode service unavailable") from exc
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=502, detail="Invalid response from pincode service") from exc

    location = _extract_postal_location(payload)
    if not location:
        raise HTTPException(status_code=404, detail="No location found for this pincode")

    district, state, country = location
    return {
        "pincode": normalized_pincode,
        "district": district,
        "state": state,
        "country": country,
    }


@app.get("/api/ifsc/{ifsc}")
def lookup_ifsc_details(ifsc: str) -> dict[str, str]:
    normalized_ifsc = re.sub(r"[^A-Za-z0-9]", "", str(ifsc or "")).upper()
    if not re.fullmatch(r"[A-Z]{4}0[A-Z0-9]{6}", normalized_ifsc):
        raise HTTPException(status_code=400, detail="IFSC format must be like BARC0INBBIR")

    lookup_url = f"https://ifsc.razorpay.com/{normalized_ifsc}"
    try:
        with urlopen(lookup_url, timeout=8) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except HTTPError as exc:
        if exc.code == 404:
            raise HTTPException(status_code=404, detail="No bank details found for this IFSC") from exc
        status_code = exc.code if isinstance(exc.code, int) and exc.code > 0 else 502
        raise HTTPException(status_code=status_code, detail="IFSC service request failed") from exc
    except URLError as exc:
        raise HTTPException(status_code=502, detail="IFSC service unavailable") from exc
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=502, detail="Invalid response from IFSC service") from exc

    details = _extract_ifsc_location(payload)
    if not details:
        raise HTTPException(status_code=404, detail="No bank details found for this IFSC")

    bank_name, branch_name = details
    return {
        "ifsc": normalized_ifsc,
        "bankName": bank_name,
        "branchName": branch_name,
    }


@app.post("/api/db/cleanup-duplicate-vendor-columns")
def db_cleanup_duplicate_vendor_columns(
    _: dict[str, Any] = Depends(require_role("Admin")),
) -> dict[str, Any]:
    return cleanup_vendor_duplicate_columns()


TEMPLATES_DIR = Path(__file__).parent / "templates"


def read_template(name: str) -> str:
    return (TEMPLATES_DIR / name).read_text(encoding="utf-8")


@app.get("/web/portal-liquid.css")
def portal_liquid_css() -> FileResponse:
    css_path = Path(__file__).resolve().parent / "assets" / "portal_liquid.css"
    if not css_path.exists():
        raise HTTPException(status_code=404, detail="Portal stylesheet not found")
    return FileResponse(
        css_path,
        media_type="text/css",
        headers={"Cache-Control": "no-store"},
    )


def _is_public_vendor_web_request(request: Request) -> bool:
    path = request.url.path
    if path not in PUBLIC_VENDOR_WEB_PATHS:
        return False
    requested_flow = str(request.query_params.get("flow") or "").strip().lower()
    return requested_flow not in {"internal", "import"}


@app.middleware("http")
async def require_login_for_web_pages(request: Request, call_next):
    path = request.url.path
    if path.startswith("/web") and path not in WEB_PUBLIC_PATHS and not _is_public_vendor_web_request(request):
        token = request.cookies.get(AUTH_COOKIE_NAME)
        session_user = user_from_token(token) if token else None
        if not token or not session_user:
            requested_path = path
            if request.url.query:
                requested_path = f"{requested_path}?{request.url.query}"
            login_url = f"/web/login?next={quote(requested_path, safe='')}"
            return RedirectResponse(url=login_url, status_code=307)
        if bool(session_user.get("passwordChangeRequired")) and path != "/web/change-password":
            requested_path = path
            if request.url.query:
                requested_path = f"{requested_path}?{request.url.query}"
            change_password_url = f"/web/change-password?next={quote(requested_path, safe='')}"
            return RedirectResponse(url=change_password_url, status_code=307)

        normalized_user_role = normalize_role(session_user.get("role"))
        bot_name = normalize_bot_name(session_user.get("botName") or "Vendor Bot")
        portal = str(request.query_params.get("portal") or "").strip().lower()

        # Admin can access both Vendor and Customer portals.
        if normalized_user_role != "admin":
            # Bot-based gating for which dashboard/form a user should see.
            # This supplements role-based access.
            if bot_name != "Both":
                is_vendor_path = (
                    path == "/web/vendor-dashboard"
                    or path == "/web/invite-vendor"
                    or (path == "/web/update-vendor" and portal != "customer")
                    or path.startswith("/web/vendor-registration")
                    or path.startswith("/web/import-vendor")
                )
                is_customer_path = (
                    path == "/web/customer-dashboard"
                    or path == "/web/invite-customer"
                    or path == "/web/customer-form"
                    or path == "/web/export-customer-form"
                    or path == "/web/customer-shipto"
                    or path == "/web/view-customer"
                    or path == "/web/update-customer"
                    or path == "/web/vendor-history"
                    or (path == "/web/view-vendor" and portal == "customer")
                    or (path == "/web/update-vendor" and portal == "customer")
                )

                if path == "/web/select-role":
                    target = (
                        "/web/customer-dashboard"
                        if bot_name == "Customer Bot"
                        else "/web/vendor-dashboard"
                    )
                    return RedirectResponse(url=target, status_code=307)

                if bot_name == "Vendor Bot" and is_customer_path:
                    return RedirectResponse(url="/web/vendor-dashboard", status_code=307)
                if bot_name == "Customer Bot" and is_vendor_path:
                    return RedirectResponse(url="/web/customer-dashboard", status_code=307)

                if path == "/web/view-vendor" and portal != "customer" and bot_name == "Customer Bot":
                    return RedirectResponse(url="/web/customer-dashboard", status_code=307)
                if path == "/web/update-vendor" and portal != "customer" and bot_name == "Customer Bot":
                    return RedirectResponse(url="/web/customer-dashboard", status_code=307)

    return await call_next(request)


@app.get("/web/login", response_class=HTMLResponse)
def web_login() -> Response:
    return HTMLResponse(content=read_template("login.html"))


@app.get("/web/change-password", response_class=HTMLResponse)
def web_change_password() -> str:
    return read_template("change_password.html")


@app.get("/web/select-role", response_class=HTMLResponse)
def web_select_role(request: Request) -> Response:
    session_user = get_cookie_session_user(request)
    bot_name = normalize_bot_name(session_user.get("botName") or "Vendor Bot")
    normalized_user_role = normalize_role(session_user.get("role"))

    # Admin must be able to choose either Vendor or Customer portal.
    if normalized_user_role == "admin":
        return HTMLResponse(content=read_template("select_role.html"))

    # If user is not allowed to choose a role, redirect them directly.
    if bot_name != "Both":
        requested_next = str(request.query_params.get("next") or "").strip()
        if requested_next.startswith("/web/"):
            return RedirectResponse(url=requested_next, status_code=307)
        default_url = (
            "/web/customer-dashboard" if bot_name == "Customer Bot" else "/web/vendor-dashboard"
        )
        return RedirectResponse(url=default_url, status_code=307)

    return HTMLResponse(content=read_template("select_role.html"))


@app.get("/web/vendor-dashboard", response_class=HTMLResponse)
def web_vendor_dashboard() -> str:
    return read_template("vendor_dashboard.html")


@app.get("/web/customer-dashboard", response_class=HTMLResponse)
def web_customer_dashboard() -> str:
    return read_template("customer_dashboard.html")


@app.get("/web/admin-users", response_class=HTMLResponse)
def web_admin_users(request: Request) -> Response:
    user = get_cookie_session_user(request)
    if normalize_role(user.get("role")) != "admin":
        return RedirectResponse(url="/web/vendor-dashboard", status_code=307)
    return HTMLResponse(content=read_template("admin_users.html"))


@app.get("/web/vendor-history", response_class=HTMLResponse)
def web_vendor_history() -> str:
    return read_template("vendor_history.html")


@app.get("/web/invite-customer", response_class=HTMLResponse)
def web_invite_customer() -> str:
    return read_template("invite_customer.html")


@app.get("/web/customer-form", response_class=HTMLResponse)
def web_customer_form() -> str:
    return read_template("customer_form.html")


@app.get("/web/export-customer-form", response_class=HTMLResponse)
def web_export_customer_form() -> str:
    return read_template("export_customer_form.html")


@app.get("/web/customer-shipto", response_class=HTMLResponse)
def web_customer_shipto() -> str:
    return read_template("customer_shipto.html")


@app.get("/web/invite-vendor", response_class=HTMLResponse)
def web_invite_vendor() -> str:
    return read_template("invite_vendor.html")


@app.get("/web/import-vendor/prescreen", response_class=HTMLResponse)
def web_import_vendor_prescreen() -> RedirectResponse:
    return RedirectResponse(url="/web/import-vendor", status_code=307)


@app.get("/web/import-vendor", response_class=HTMLResponse)
def web_import_vendor() -> str:
    return read_template("import_vendor.html")


@app.get("/web/vendor-registration", response_class=HTMLResponse)
def web_vendor_registration(request: Request):
    requested_flow = str(request.query_params.get("flow") or "domestic").strip().lower()
    if requested_flow == "import":
        return RedirectResponse(url="/web/import-vendor", status_code=307)
    if requested_flow == "internal":
        return RedirectResponse(url="/web/internal-registration/details", status_code=307)
    return read_template("vendor_registration_entry.html")


@app.get("/web/vendor-registration/prescreen", response_class=HTMLResponse)
def web_vendor_registration_prescreen() -> str:
    return read_template("vendor_prescreen.html")


@app.get("/web/vendor-registration/details", response_class=HTMLResponse)
def web_vendor_registration_details() -> str:
    return read_template("vendor_registration.html")


@app.get("/web/internal-registration", response_class=HTMLResponse)
def web_internal_registration() -> RedirectResponse:
    return RedirectResponse(url="/web/internal-registration/details", status_code=307)


@app.get("/web/internal-registration/details", response_class=HTMLResponse)
def web_internal_registration_details() -> str:
    return read_template("internal_registration.html")


@app.get("/web/view-vendor", response_class=HTMLResponse)
def web_view_vendor() -> str:
    return read_template("view_vendor.html")


@app.get("/web/view-customer", response_class=HTMLResponse)
def web_view_customer() -> str:
    return read_template("view_customer.html")


@app.get("/web/update-vendor", response_class=HTMLResponse)
def web_update_vendor() -> str:
    return read_template("update_vendor.html")


@app.get("/web/update-customer", response_class=HTMLResponse)
def web_update_customer() -> str:
    return read_template("update_customer.html")


@app.get("/web/code-of-conduct-pdf")
def code_of_conduct_pdf() -> Response:
    try:
        pdf_bytes = _build_code_of_conduct_pdf_bytes()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Unable to render Code of Conduct document: {exc}") from exc
    return Response(content=pdf_bytes, media_type="application/pdf")


@app.get("/web/code-of-conduct-docx")
def code_of_conduct_docx() -> FileResponse:
    if not CODE_OF_CONDUCT_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=404, detail="Code of Conduct document not found")
    return FileResponse(
        CODE_OF_CONDUCT_TEMPLATE_PATH,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=CODE_OF_CONDUCT_TEMPLATE_PATH.name,
    )


@app.get("/web/code-of-conduct-preview", response_class=HTMLResponse)
def code_of_conduct_preview() -> HTMLResponse:
    try:
        html_content = _build_code_of_conduct_view_html()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Unable to render Code of Conduct preview: {exc}") from exc
    return HTMLResponse(content=html_content, headers={"Cache-Control": "no-store"})


@app.get("/web/code-of-conduct-viewer", response_class=HTMLResponse)
def code_of_conduct_viewer() -> HTMLResponse:
    try:
        html_content = _build_code_of_conduct_view_html()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Unable to render Code of Conduct viewer: {exc}") from exc
    return HTMLResponse(content=html_content, headers={"Cache-Control": "no-store"})


@app.get("/web/logo-image")
def logo_image() -> FileResponse:
    logo_path = Path(__file__).resolve().parent / "assets" / "branding" / "sharp_tannan_logo.png"
    if not logo_path.exists():
        logo_path = Path(__file__).resolve().parents[2] / "front" / "public" / "logo.png"
    if not logo_path.exists():
        raise HTTPException(status_code=404, detail="Logo image not found")
    return FileResponse(logo_path, media_type="image/png")


@app.get("/web/header-art-image")
def header_art_image() -> FileResponse:
    image_path = Path(__file__).resolve().parent / "assets" / "branding" / "sharp_tannan_logo.png"
    if not image_path.exists():
        image_path = Path(__file__).resolve().parents[2] / "front" / "public" / "download-artguru.png"
    if not image_path.exists():
        raise HTTPException(status_code=404, detail="Header art image not found")
    return FileResponse(image_path, media_type="image/png")


@app.get("/mock-files/{file_name:path}")
def get_mock_file(file_name: str) -> Response:
    storage_key, normalized_name = _normalize_mock_file_reference(file_name)
    if not normalized_name:
        raise HTTPException(status_code=404, detail="File not found")

    stored = mock_files.get(storage_key) or mock_files.get(normalized_name)
    if not stored:
        disk_path = _resolve_uploaded_file_disk_path(storage_key)
        if disk_path is not None and disk_path.exists() and disk_path.is_file():
            return _response_from_uploaded_disk_file(disk_path)

        disk_path = _find_uploaded_file_by_name(normalized_name)
        if disk_path is not None:
            return _response_from_uploaded_disk_file(disk_path)

        stored = get_vendor_document_content_by_filename(normalized_name)
        if not stored:
            raise HTTPException(status_code=404, detail="File not found")

    stored_content = bytes(stored.get("content") or b"")
    decoded_reference = _decode_document_reference_text_payload(stored_content)
    if decoded_reference:
        lower_reference = decoded_reference.lower()
        if lower_reference.startswith("http://") or lower_reference.startswith("https://"):
            proxied_remote_response = _build_inline_document_response_from_remote_url(decoded_reference)
            if proxied_remote_response is not None:
                return proxied_remote_response
            return RedirectResponse(url=decoded_reference)
        if decoded_reference.startswith("/mock-files/"):
            nested_reference = decoded_reference.split("/mock-files/", 1)[1]
            nested_normalized_name = _extract_document_filename(nested_reference)
            if nested_normalized_name and nested_normalized_name.lower() != normalized_name.lower():
                try:
                    return get_mock_file(nested_reference)
                except HTTPException as exc:
                    if exc.status_code != 404:
                        raise
        elif decoded_reference.lower() != normalized_name.lower():
            try:
                return get_mock_file(decoded_reference)
            except HTTPException as exc:
                if exc.status_code != 404:
                    raise

    if normalized_name.lower().endswith(".pdf") and not _has_valid_document_signature(normalized_name, stored_content):
        raise HTTPException(status_code=404, detail="File not found")

    return _build_inline_document_response(
        normalized_name,
        stored_content,
        stored.get("content_type"),
    )


@app.get("/api/vendor/document")
def get_vendor_document(
    recordId: str,
    fieldName: str,
    user: dict[str, Any] | None = Depends(get_optional_current_user),
) -> Response:
    record_id = _normalize_record_id(recordId)
    field_name = str(fieldName or "").strip()
    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")
    if not field_name:
        raise HTTPException(status_code=400, detail="fieldName is required")

    local_record = find_local_vendor_record_by_record_id(record_id)
    db_record = get_vendor_record_by_record_id(record_id)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    _authorize_vendor_record_access(local_record, db_record, user)

    candidate_keys = _get_document_candidate_keys(field_name)
    canonical_field_name = candidate_keys[0] if candidate_keys else field_name

    stored_reference_hint = ""
    stored_document = get_vendor_document_content_by_record_id(record_id, candidate_keys)
    if stored_document:
        file_name = _extract_document_filename(stored_document.get("filename")) or "document"
        stored_content = bytes(stored_document.get("content") or b"")
        decoded_reference = _decode_document_reference_text_payload(stored_content)
        if decoded_reference:
            stored_reference_hint = decoded_reference
        elif _has_valid_document_signature(file_name, stored_content):
            return _build_inline_document_response(
                file_name,
                stored_content,
                stored_document.get("content_type"),
            )

    raw_reference_candidates: list[str] = []
    seen_references: set[str] = set()

    if stored_reference_hint:
        marker = stored_reference_hint.lower()
        if marker not in seen_references:
            seen_references.add(marker)
            raw_reference_candidates.append(stored_reference_hint)

    for record in (db_record, local_record):
        if not isinstance(record, dict):
            continue
        for key in candidate_keys:
            value = str(record.get(key) or "").strip()
            if not value:
                continue
            dedupe_key = value.lower()
            if dedupe_key in seen_references:
                continue
            seen_references.add(dedupe_key)
            raw_reference_candidates.append(value)

    extracted_file_names: list[str] = []
    seen_file_names: set[str] = set()
    for raw_reference in raw_reference_candidates:
        file_name = _extract_document_filename(raw_reference)
        if not file_name:
            continue
        dedupe_key = file_name.lower()
        if dedupe_key in seen_file_names:
            continue
        seen_file_names.add(dedupe_key)
        extracted_file_names.append(file_name)

    disk_path: Path | None = None
    field_name_candidates = [canonical_field_name, *candidate_keys[1:]]
    for field_key in field_name_candidates:
        for extracted_file_name in extracted_file_names:
            disk_path = _find_uploaded_file_for_record_field(record_id, field_key, extracted_file_name)
            if disk_path is not None:
                break
        if disk_path is not None:
            break
        disk_path = _find_uploaded_file_for_record_field(record_id, field_key)
        if disk_path is not None:
            break
    if disk_path is not None:
        return _response_from_uploaded_disk_file(disk_path)

    for extracted_file_name in extracted_file_names:
        disk_path = _find_uploaded_file_by_name(extracted_file_name)
        if disk_path is not None:
            return _response_from_uploaded_disk_file(disk_path)

        stored = get_vendor_document_content_by_filename(extracted_file_name)
        if stored:
            stored_content = bytes(stored.get("content") or b"")
            decoded_reference = _decode_document_reference_text_payload(stored_content)
            if decoded_reference:
                marker = decoded_reference.lower()
                if marker not in seen_references:
                    seen_references.add(marker)
                    raw_reference_candidates.append(decoded_reference)
            if not _has_valid_document_signature(extracted_file_name, stored_content):
                continue
            return _build_inline_document_response(
                extracted_file_name,
                stored_content,
                stored.get("content_type"),
            )

    for raw_reference in raw_reference_candidates:
        mock_reference = raw_reference
        lower_reference = mock_reference.lower()
        if lower_reference.startswith("http://") or lower_reference.startswith("https://"):
            proxied_remote_response = _build_inline_document_response_from_remote_url(mock_reference)
            if proxied_remote_response is not None:
                return proxied_remote_response
            return RedirectResponse(url=mock_reference)
        if mock_reference.startswith("/mock-files/"):
            mock_reference = mock_reference.split("/mock-files/", 1)[1]
        try:
            return get_mock_file(mock_reference)
        except HTTPException as exc:
            if exc.status_code != 404:
                raise

    print(
        "[vendor-document] file not found",
        {
            "recordId": record_id,
            "fieldName": field_name,
            "candidateKeys": candidate_keys,
            "rawReferenceCandidates": [
                str(value)[:250]
                for value in raw_reference_candidates[:10]
            ],
            "extractedFileNames": extracted_file_names[:10],
        },
    )
    raise HTTPException(status_code=404, detail="File not found")


@app.get("/api/vendor/document/pdf")
def get_vendor_document_pdf(
    recordId: str,
    fieldName: str,
    user: dict[str, Any] | None = Depends(get_optional_current_user),
) -> Response:
    document_response = get_vendor_document(recordId=recordId, fieldName=fieldName, user=user)
    return _build_pdf_only_response_from_document_response(document_response, fallback_file_name="document.pdf")


@app.post("/api/auth/register")
def register(body: dict[str, Any]) -> dict[str, Any]:
    name = empty_to_none(body.get("name"))
    email = empty_to_none(body.get("email"))
    password = empty_to_none(body.get("password"))
    role = empty_to_none(body.get("role")) or "Vendor"

    if not name or not email or not password:
        raise HTTPException(status_code=400, detail="Name, email and password are required")

    email_key = str(email).lower()
    if email_key in users_by_email:
        raise HTTPException(status_code=409, detail="User with this email already exists")

    create_user(str(name), str(email), str(password), str(role))
    return {"message": "User registered successfully"}


@app.post("/api/auth/login")
def login(body: dict[str, Any], response: Response) -> dict[str, Any]:
    login_id = empty_to_none(body.get("email") or body.get("username") or body.get("loginId"))
    password = empty_to_none(body.get("password"))

    if not login_id or not password:
        raise HTTPException(status_code=400, detail="Email/Username and password are required")

    login_id_str = str(login_id).strip()
    password_str = str(password)

    # First preference: SQL Users table auth.
    db_user = authenticate_db_user(login_id_str, password_str)
    if db_user:
        user = ensure_local_user(
            name=str(db_user["name"]),
            email=str(db_user["email"]),
            role=str(db_user["role"]),
            preferred_id=db_user.get("id"),
            bot_name=str(db_user.get("botName") or "Vendor Bot").strip() or "Vendor Bot",
        )
    else:
        db_info = get_db_status()
        if db_info.get("enabled") and not db_info.get("connected"):
            raise HTTPException(
                status_code=503,
                detail=f"Database is not connected. {db_info.get('error') or ''}".strip(),
            )
        uid = users_by_email.get(login_id_str.lower())
        user = users.get(uid or -1)
        if not user:
            login_lower = login_id_str.lower()
            for local_user in users.values():
                if str(local_user.get("name") or "").strip().lower() == login_lower:
                    user = local_user
                    break
        if not user or not check_password(password_str, user["password"]):
            raise HTTPException(status_code=401, detail="Invalid email/username or password")

    password_change_required = password_str == DEFAULT_NEW_USER_PASSWORD
    user["passwordChangeRequired"] = password_change_required
    user.setdefault("botName", "Vendor Bot")

    token = make_token(user, password_change_required=password_change_required)
    response.set_cookie(
        key=AUTH_COOKIE_NAME,
        value=token,
        httponly=True,
        samesite="lax",
        secure=False,
        path="/",
    )
    return {
        "message": "Login successful",
        "token": token,
        "role": user["role"],
        "botName": normalize_bot_name(user.get("botName") or "Vendor Bot"),
        "email": user["email"],
        "name": user["name"],
        "passwordChangeRequired": password_change_required,
    }


@app.post("/api/auth/forgot-password")
def forgot_password(body: dict[str, Any]) -> dict[str, Any]:
    email = str(empty_to_none(body.get("email")) or "").strip()
    new_password = str(empty_to_none(body.get("newPassword") or body.get("password")) or "")

    if not email or not new_password:
        raise HTTPException(status_code=400, detail="email and newPassword are required")
    if "@" not in email or "." not in email.split("@", 1)[-1]:
        raise HTTPException(status_code=400, detail="Invalid email format")
    if len(new_password) < 6:
        raise HTTPException(status_code=400, detail="Password must be at least 6 characters")

    email_key = email.lower()
    local_updated = False
    local_user_id = users_by_email.get(email_key)
    if local_user_id is not None and local_user_id in users:
        users[local_user_id]["password"] = hash_password(new_password)
        users[local_user_id]["passwordChangeRequired"] = False
        local_updated = True

    db_updated = False
    db_update_error: str | None = None
    try:
        db_updated = update_db_user_password_by_email(email, new_password)
    except RuntimeError as exc:
        db_update_error = str(exc)

    if not local_updated and not db_updated:
        if db_update_error:
            raise HTTPException(status_code=500, detail=db_update_error)
        raise HTTPException(status_code=404, detail="Email ID not found")

    return {"message": "Password updated successfully"}


@app.post("/api/auth/change-password")
def change_password(
    body: dict[str, Any],
    response: Response,
    current_user: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    current_password = str(empty_to_none(body.get("currentPassword")) or "")
    new_password = str(empty_to_none(body.get("newPassword") or body.get("password")) or "")

    if not current_password or not new_password:
        raise HTTPException(status_code=400, detail="currentPassword and newPassword are required")
    if len(new_password) < 6:
        raise HTTPException(status_code=400, detail="Password must be at least 6 characters")
    if current_password == new_password:
        raise HTTPException(status_code=400, detail="New password must be different from current password")

    email = str(current_user.get("email") or "").strip()
    if not email:
        raise HTTPException(status_code=400, detail="User email not found")
    email_key = email.lower()

    local_user_id = users_by_email.get(email_key)
    local_user = users.get(local_user_id or -1)
    local_password_matches = False
    if local_user and str(local_user.get("password") or "").strip():
        local_password_matches = check_password(current_password, str(local_user.get("password") or ""))

    db_password_matches = False
    db_user = authenticate_db_user(email, current_password)
    if db_user and str(db_user.get("email") or "").strip().lower() == email_key:
        db_password_matches = True

    if not local_password_matches and not db_password_matches:
        raise HTTPException(status_code=401, detail="Current password is incorrect")

    if not local_user:
        local_user = ensure_local_user(
            name=str(current_user.get("name") or email).strip(),
            email=email,
            role=str(current_user.get("role") or "User").strip() or "User",
            preferred_id=current_user.get("id"),
        )

    local_user["password"] = hash_password(new_password)
    local_user["passwordChangeRequired"] = False

    if db_password_matches:
        try:
            db_updated = update_db_user_password_by_email(email, new_password)
        except RuntimeError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        if not db_updated:
            raise HTTPException(status_code=500, detail="Unable to update password in database")

    new_token = make_token(local_user, password_change_required=False)
    response.set_cookie(
        key=AUTH_COOKIE_NAME,
        value=new_token,
        httponly=True,
        samesite="lax",
        secure=False,
        path="/",
    )
    return {
        "message": "Password changed successfully",
        "token": new_token,
        "role": local_user.get("role"),
        "email": local_user.get("email"),
        "name": local_user.get("name"),
        "passwordChangeRequired": False,
    }


@app.post("/api/auth/logout")
def logout(response: Response) -> dict[str, str]:
    response.delete_cookie(key=AUTH_COOKIE_NAME, path="/")
    return {"message": "Logout successful"}


ADMIN_ALLOWED_ROLE_LABELS: dict[str, str] = {
    "user": "User",
    "hod": "HOD",
    "validator": "Validator",
    "validator & user": "Validator & User",
    "validator and user": "Validator & User",
    "validator+user": "Validator & User",
    "validator_user": "Validator & User",
    "validator & hod": "Validator & HOD",
    "validator and hod": "Validator & HOD",
    "validator+hod": "Validator & HOD",
    "validator_hod": "Validator & HOD",
}
ADMIN_ALLOWED_BOT_LABELS: dict[str, str] = {
    "vendor bot": "Vendor Bot",
    "customer bot": "Customer Bot",
    "both": "Both",
}


def _resolve_admin_user_role_label(role_input: Any) -> str:
    raw_role = str(empty_to_none(role_input) or "").strip()
    normalized_role = normalize_role(raw_role)
    return ADMIN_ALLOWED_ROLE_LABELS.get(normalized_role, "")


def _resolve_admin_bot_name(bot_input: Any) -> str:
    raw_bot = str(empty_to_none(bot_input) or "").strip()
    normalized_bot = str(raw_bot or "").strip().lower()
    return ADMIN_ALLOWED_BOT_LABELS.get(normalized_bot, "Vendor Bot")


@app.get("/api/admin/users")
def admin_get_users(_: dict[str, Any] = Depends(require_role("Admin"))) -> list[dict[str, Any]]:
    rows = get_db_users()
    for row in rows:
        email = str(row.get("email") or "").strip()
        if not email:
            continue
        ensure_local_user(
            name=str(row.get("name") or email).strip(),
            email=email,
            role=str(row.get("role") or "User").strip() or "User",
            preferred_id=row.get("id"),
            bot_name=str(row.get("botName") or "Vendor Bot").strip(),
        )
    return rows


@app.post("/api/admin/users")
def admin_create_user(
    body: dict[str, Any],
    request: Request,
    _: dict[str, Any] = Depends(require_role("Admin")),
) -> dict[str, Any]:
    name = str(empty_to_none(body.get("name")) or "").strip()
    email = str(empty_to_none(body.get("email")) or "").strip()
    password = DEFAULT_NEW_USER_PASSWORD
    role = _resolve_admin_user_role_label(body.get("role") or "User")
    bot_name = _resolve_admin_bot_name(body.get("botName") or body.get("bot_name") or "Vendor Bot")

    if not name or not email:
        raise HTTPException(status_code=400, detail="Name and email are required")
    if "@" not in email or "." not in email.split("@", 1)[-1]:
        raise HTTPException(status_code=400, detail="Invalid email format")
    if not role:
        raise HTTPException(
            status_code=400,
            detail="Role must be User, HOD, Validator, Validator & User, or Validator & HOD",
        )

    try:
        created = insert_db_user(
            name=name,
            email=email,
            password=password,
            role=role,
            bot_name=bot_name,
        )
    except ValueError as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    local_user = ensure_local_user(
        name=str(created.get("name") or name).strip(),
        email=str(created.get("email") or email).strip(),
        role=str(created.get("role") or role).strip() or "User",
        preferred_id=created.get("id"),
        bot_name=bot_name,
    )
    local_user["password"] = hash_password(password)
    local_user["passwordChangeRequired"] = True

    welcome_email_sent = False
    welcome_email_error: str | None = None
    login_url = _build_public_registration_url(request, "/web/login")
    try:
        send_new_user_welcome_email(
            recipient_name=local_user["name"],
            recipient_email=local_user["email"],
            username=local_user["email"],
            password=password,
            login_url=login_url,
        )
        welcome_email_sent = True
    except Exception as exc:
        welcome_email_error = str(exc)

    return {
        "message": "User created successfully",
        "defaultPassword": password,
        "welcomeEmailSent": welcome_email_sent,
        "welcomeEmailError": welcome_email_error,
        "user": {
            "id": local_user["id"],
            "name": local_user["name"],
            "email": local_user["email"],
            "role": local_user["role"],
            "botName": local_user.get("botName") or bot_name,
        },
    }


@app.put("/api/admin/users")
def admin_update_user(
    body: dict[str, Any],
    _: dict[str, Any] = Depends(require_role("Admin")),
) -> dict[str, Any]:
    current_email = str(empty_to_none(body.get("currentEmail")) or "").strip()
    name = str(empty_to_none(body.get("name")) or "").strip()
    email = str(empty_to_none(body.get("email")) or "").strip()
    role = _resolve_admin_user_role_label(body.get("role") or "User")
    bot_name = _resolve_admin_bot_name(body.get("botName") or body.get("bot_name") or "Vendor Bot")

    if not current_email:
        raise HTTPException(status_code=400, detail="currentEmail is required")
    if not name or not email:
        raise HTTPException(status_code=400, detail="Name and email are required")
    if "@" not in email or "." not in email.split("@", 1)[-1]:
        raise HTTPException(status_code=400, detail="Invalid email format")
    if not role:
        raise HTTPException(
            status_code=400,
            detail="Role must be User, HOD, Validator, Validator & User, or Validator & HOD",
        )

    try:
        updated = update_db_user_by_email(
            current_email=current_email,
            name=name,
            email=email,
            role=role,
            bot_name=bot_name,
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    if not updated:
        raise HTTPException(status_code=404, detail="User not found")

    previous_email_key = current_email.lower()
    updated_email = str(updated.get("email") or email).strip()
    updated_name = str(updated.get("name") or name).strip()
    updated_role = str(updated.get("role") or role).strip() or "User"
    updated_bot_name = str(updated.get("botName") or bot_name).strip() or "Vendor Bot"
    local_user_id = users_by_email.get(previous_email_key)
    if local_user_id is not None and local_user_id in users:
        local_user = users[local_user_id]
        if previous_email_key != updated_email.lower():
            users_by_email.pop(previous_email_key, None)
            local_user["email"] = updated_email
            users_by_email[updated_email.lower()] = local_user_id
        local_user["name"] = updated_name
        local_user["role"] = updated_role
        local_user["botName"] = normalize_bot_name(updated_bot_name)
    else:
        ensure_local_user(
            name=updated_name,
            email=updated_email,
            role=updated_role,
            preferred_id=updated.get("id"),
            bot_name=updated_bot_name,
        )

    return {
        "message": "User updated successfully",
        "user": {
            "id": updated.get("id"),
            "name": updated_name,
            "email": updated_email,
            "role": updated_role,
            "botName": normalize_bot_name(updated_bot_name),
        },
    }


@app.get("/api/vendor/get-buyer-emails")
def get_buyer_emails() -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}

    for user in users.values():
        email = str(user.get("email") or "").strip()
        if user_has_any_role(user, "user", "buyer") and email:
            merged[email.lower()] = {
                "id": user.get("id"),
                "name": str(user.get("name") or email).strip(),
                "email": email,
            }

    for row in get_db_users():
        email = str(row.get("email") or "").strip()
        if not email:
            continue
        if not any(role in get_role_capabilities(row.get("role")) for role in {"user", "buyer"}):
            continue
        role_value = str(row.get("role") or "User").strip() or "User"
        local_user = ensure_local_user(
            name=str(row.get("name") or email).strip(),
            email=email,
            role=role_value,
            preferred_id=row.get("id"),
            bot_name=str(row.get("botName") or "Vendor Bot").strip() or "Vendor Bot",
        )
        merged[email.lower()] = {
            "id": local_user["id"],
            "name": local_user["name"],
            "email": email,
        }

    values = list(merged.values())
    values.sort(key=lambda item: (str(item.get("name") or "").lower(), str(item.get("email") or "").lower()))
    return values


@app.get("/api/vendor/get-hod-emails")
def get_vendor_hod_emails() -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    hod_role_keys = {"hod", "validator_hod"}

    for user in users.values():
        role_key = normalize_role(user.get("role"))
        email = str(user.get("email") or "").strip()
        if role_key in hod_role_keys and email:
            merged[email.lower()] = {"id": user.get("id"), "name": user.get("name") or email, "email": email}

    for row in get_db_users():
        role_key = normalize_role(row.get("role"))
        email = str(row.get("email") or "").strip()
        if not email:
            continue
        if role_key not in hod_role_keys:
            continue
        merged[email.lower()] = {
            "id": row.get("id"),
            "name": str(row.get("name") or email).strip(),
            "email": email,
        }

    values = list(merged.values())
    values.sort(key=lambda item: (str(item.get("name") or "").lower(), str(item.get("email") or "").lower()))
    return values


def _get_validator_contacts() -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}

    for user in users.values():
        email = str(user.get("email") or "").strip()
        if not email or not user_has_any_role(user, "validator"):
            continue
        merged[email.lower()] = {
            "id": user.get("id"),
            "name": str(user.get("name") or email).strip(),
            "email": email,
        }

    for row in get_db_users():
        email = str(row.get("email") or "").strip()
        if not email:
            continue
        if "validator" not in get_role_capabilities(row.get("role")):
            continue
        role_value = str(row.get("role") or "Validator").strip() or "Validator"
        local_user = ensure_local_user(
            name=str(row.get("name") or email).strip(),
            email=email,
            role=role_value,
            preferred_id=row.get("id"),
            bot_name=str(row.get("botName") or "Vendor Bot").strip() or "Vendor Bot",
        )
        merged[email.lower()] = {
            "id": local_user.get("id"),
            "name": str(local_user.get("name") or email).strip(),
            "email": email,
        }

    values = list(merged.values())
    values.sort(key=lambda item: (str(item.get("name") or "").lower(), str(item.get("email") or "").lower()))
    return values


def _pick_first_record_text_value(record: dict[str, Any] | None, *keys: str) -> str:
    if not isinstance(record, dict):
        return ""
    for key in keys:
        value = record.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def _extract_vendor_status_text(record: dict[str, Any] | None, default_status: str = "") -> str:
    return _pick_first_record_text_value(
        record,
        "approverStatus",
        "ApproverStatus",
        "ApprovStatus",
        "inviteStatus",
        "status",
        "Status",
    ) or str(default_status or "").strip()


def _is_created_in_sap_record(record: dict[str, Any] | None) -> bool:
    if not isinstance(record, dict):
        return False

    vendor_code = _pick_first_record_text_value(
        record,
        "vendorCode",
        "VendorCode",
        "SAPVendorCode",
        "BusinessPartnerCode",
        "sapVendorCode",
        "businessPartnerCode",
    )
    sap_code_generated_at = _pick_first_record_text_value(
        record,
        "SAPCodeGeneratedAt",
        "sapCodeGeneratedAt",
    )
    if vendor_code or sap_code_generated_at:
        return True

    normalized_status = _normalize_workflow_status_key(_extract_vendor_status_text(record))
    return (
        "created in sap" in normalized_status
        or "sap code created" in normalized_status
        or "vendor code created" in normalized_status
    )


def _map_vendor_summary_row(record: dict[str, Any] | None, default_status: str = "") -> dict[str, Any]:
    return {
        "record": _pick_first_record_text_value(record, "recordId", "RecordID", "record", "Record", "id", "Id", "ID"),
        "company": _pick_first_record_text_value(record, "companyName", "CompanyName", "company", "Company"),
        "category": _pick_first_record_text_value(record, "vendorCategory", "VendorCategory", "Vendor_Category", "category", "Category"),
        "approverStatus": _extract_vendor_status_text(record, default_status=default_status),
        "name": _pick_first_record_text_value(record, "vendorName", "Vendorname", "VendorName", "name", "Name"),
        "type": _pick_first_record_text_value(record, "vendorType", "VendorType", "type", "Type"),
        "vendorCode": _pick_first_record_text_value(
            record,
            "vendorCode",
            "VendorCode",
            "SAPVendorCode",
            "BusinessPartnerCode",
            "sapVendorCode",
            "businessPartnerCode",
        ),
        "sapCodeGeneratedAt": _pick_first_record_text_value(
            record,
            "SAPCodeGeneratedAt",
            "sapCodeGeneratedAt",
        ),
        "createdAt": _pick_first_record_text_value(
            record,
            "createdAt",
            "CreatedAt",
            "created_on",
            "CreatedOn",
            "updatedAt",
            "UpdatedAt",
        ),
    }


def _is_pending_user_approval_status_value(status_value: Any) -> bool:
    normalized = " ".join(str(status_value or "").strip().lower().split())
    if not normalized:
        return False
    known_pending_user_statuses = {
        "pending for user approval",
        "pendign for user approval",
        "pending for buyer approval",
        "pendign for buyer approval",
        "pending for data validation",
    }
    if normalized in known_pending_user_statuses:
        return True
    return "pending" in normalized and ("user approv" in normalized or "buyer approv" in normalized)


@app.get("/api/vendor/dashboard-records")
def get_dashboard_records(user: dict[str, Any] = Depends(get_current_user)) -> list[dict[str, Any]]:
    normalized_role = normalize_role(user.get("role"))

    if normalized_role in {"validator_user", "validator_hod"}:
        merged_rows: dict[str, dict[str, Any]] = {}
        role_sequence = (
            ("validator", "user")
            if normalized_role == "validator_user"
            else ("validator", "hod")
        )
        for role_key in role_sequence:
            scoped_user = dict(user)
            scoped_user["role"] = role_key
            for row in get_dashboard_records(scoped_user):
                record_key = str(row.get("record") or "").strip().lower()
                if not record_key:
                    record_key = f"temp-{len(merged_rows) + 1}"
                existing_row = merged_rows.get(record_key)
                if not existing_row or str(row.get("createdAt") or "") >= str(existing_row.get("createdAt") or ""):
                    merged_rows[record_key] = row
        result_rows = list(merged_rows.values())
        result_rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
        return result_rows

    merged: dict[str, dict[str, Any]] = {}

    def upsert_row(row: dict[str, Any]) -> None:
        key = str(row.get("record") or "").strip().lower()
        if not key:
            key = f"temp-{len(merged) + 1}"
        existing = merged.get(key)
        if not existing:
            merged[key] = row
            return
        if str(row.get("createdAt") or "") >= str(existing.get("createdAt") or ""):
            merged[key] = row

    def build_rows() -> list[dict[str, Any]]:
        rows = list(merged.values())
        rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
        return rows

    sql_rows = get_all_vendor_records()

    if normalized_role == "admin":
        for invite in invite_forms.values():
            upsert_row(_map_vendor_summary_row(invite, default_status="Vendor Invited"))

        for form in vendor_forms.values():
            upsert_row(_map_vendor_summary_row(form, default_status="Pending For User Approval"))

        for db_row in sql_rows:
            upsert_row(_map_vendor_summary_row(db_row))

        return build_rows()

    if normalized_role == "hod":
        hod_email = str(user.get("email") or "").strip().lower()
        if not hod_email:
            return []

        pending_status_keys = {
            "pending for hod approval",
            "pendign for hod approval",
        }

        def is_pending_for_hod(status_value: Any) -> bool:
            normalized_status = re.sub(r"\s+", " ", str(status_value or "").strip().lower())
            if normalized_status in pending_status_keys:
                return True
            return (
                "pending" in normalized_status
                and "hod" in normalized_status
                and "approv" in normalized_status
            )

        def is_assigned_to_hod(record: dict[str, Any] | None) -> bool:
            return _extract_record_hod_email(record) == hod_email

        for invite in invite_forms.values():
            status_value = invite.get("approverStatus") or invite.get("status")
            if not is_pending_for_hod(status_value):
                continue
            if not is_assigned_to_hod(invite):
                continue
            upsert_row(
                {
                    "record": str(invite.get("recordId") or ""),
                    "company": str(invite.get("companyName") or ""),
                    "category": str(invite.get("vendorCategory") or ""),
                    "approverStatus": str(invite.get("approverStatus") or "Pending For HOD Approval"),
                    "name": str(invite.get("vendorName") or ""),
                    "type": str(invite.get("vendorType") or ""),
                    "createdAt": str(invite.get("createdAt") or ""),
                }
            )

        for form in vendor_forms.values():
            if not is_pending_for_hod(form.get("approverStatus")):
                continue
            if not is_assigned_to_hod(form):
                continue
            upsert_row(
                {
                    "record": str(form.get("recordId") or form.get("id") or ""),
                    "company": str(form.get("companyName") or ""),
                    "category": str(form.get("vendorCategory") or ""),
                    "approverStatus": str(form.get("approverStatus") or "Pending For HOD Approval"),
                    "name": str(form.get("vendorName") or ""),
                    "type": str(form.get("vendorType") or ""),
                    "createdAt": str(form.get("createdAt") or ""),
                }
            )

        # Use full SQL rows here so HOD records still appear even when status/HOD
        # fields are stored under alternate column spellings.
        for db_row in sql_rows:
            if not is_pending_for_hod(_extract_vendor_status_text(db_row)):
                continue
            if not is_assigned_to_hod(db_row):
                continue
            summary_row = _map_vendor_summary_row(db_row, default_status="Pending For HOD Approval")
            if not str(summary_row.get("record") or "").strip():
                continue
            upsert_row(summary_row)

        for db_row in sql_rows:
            if not _is_created_in_sap_record(db_row):
                continue
            if not is_assigned_to_hod(db_row):
                continue
            summary_row = _map_vendor_summary_row(db_row)
            if not str(summary_row.get("record") or "").strip():
                continue
            upsert_row(summary_row)

        return build_rows()

    if normalized_role == "validator":
        pending_status_keys = {
            "pending for validator approval",
            "pendign for validator approval",
        }

        def is_pending_for_validator(status_value: Any) -> bool:
            return str(status_value or "").strip().lower() in pending_status_keys

        for invite in invite_forms.values():
            status_value = invite.get("approverStatus") or invite.get("status")
            if not is_pending_for_validator(status_value):
                continue
            upsert_row(
                {
                    "record": str(invite.get("recordId") or ""),
                    "company": str(invite.get("companyName") or ""),
                    "category": str(invite.get("vendorCategory") or ""),
                    "approverStatus": str(invite.get("approverStatus") or "Pending For Validator Approval"),
                    "name": str(invite.get("vendorName") or ""),
                    "type": str(invite.get("vendorType") or ""),
                    "createdAt": str(invite.get("createdAt") or ""),
                }
            )

        for form in vendor_forms.values():
            if not is_pending_for_validator(form.get("approverStatus")):
                continue
            upsert_row(
                {
                    "record": str(form.get("recordId") or form.get("id") or ""),
                    "company": str(form.get("companyName") or ""),
                    "category": str(form.get("vendorCategory") or ""),
                    "approverStatus": str(form.get("approverStatus") or "Pending For Validator Approval"),
                    "name": str(form.get("vendorName") or ""),
                    "type": str(form.get("vendorType") or ""),
                    "createdAt": str(form.get("createdAt") or ""),
                }
            )

        for status_value in ("Pending For Validator Approval", "Pendign For Validator Approval"):
            for db_row in get_vendor_records_by_approver_status(status_value):
                upsert_row(
                    {
                        "record": str(db_row.get("record") or ""),
                        "company": str(db_row.get("company") or ""),
                        "category": str(db_row.get("category") or ""),
                        "approverStatus": str(db_row.get("approverStatus") or ""),
                        "name": str(db_row.get("name") or ""),
                        "type": str(db_row.get("type") or ""),
                        "createdAt": str(db_row.get("createdAt") or ""),
                    }
                )

        for db_row in sql_rows:
            if not _is_created_in_sap_record(db_row):
                continue
            summary_row = _map_vendor_summary_row(db_row)
            if not str(summary_row.get("record") or "").strip():
                continue
            upsert_row(summary_row)

        return build_rows()

    user_email = str(user.get("email") or "").strip().lower()
    if not user_email:
        return []

    for invite in invite_forms.values():
        if not _record_is_assigned_to_email(invite, user_email):
            continue
        upsert_row(
            {
                "record": str(invite.get("recordId") or ""),
                "company": str(invite.get("companyName") or ""),
                "category": str(invite.get("vendorCategory") or ""),
                "approverStatus": str(invite.get("approverStatus") or "Vendor Invited"),
                "name": str(invite.get("vendorName") or ""),
                "type": str(invite.get("vendorType") or ""),
                "createdAt": str(invite.get("createdAt") or ""),
            }
        )

    for form in vendor_forms.values():
        if not _record_is_assigned_to_email(form, user_email):
            continue
        upsert_row(
            {
                "record": str(form.get("recordId") or form.get("id") or ""),
                "company": str(form.get("companyName") or ""),
                "category": str(form.get("vendorCategory") or ""),
                "approverStatus": str(form.get("approverStatus") or "Pending For User Approval"),
                "name": str(form.get("vendorName") or ""),
                "type": str(form.get("vendorType") or ""),
                "createdAt": str(form.get("createdAt") or ""),
            }
        )

    for db_row in get_vendor_records_by_assignee_email(user_email):
        upsert_row(
            {
                "record": str(db_row.get("record") or ""),
                "company": str(db_row.get("company") or ""),
                "category": str(db_row.get("category") or ""),
                "approverStatus": str(db_row.get("approverStatus") or ""),
                "name": str(db_row.get("name") or ""),
                "type": str(db_row.get("type") or ""),
                "createdAt": str(db_row.get("createdAt") or ""),
            }
        )

    # Fallback: scan pending-user statuses from DB and re-check assignment
    # using full-record email extraction. This handles records where assignee
    # email is stored in text/mixed columns (for example, "Name <email>").
    pending_user_statuses = (
        "Pending For User Approval",
        "Pendign For User Approval",
        "Pending For Buyer Approval",
        "Pendign For Buyer Approval",
        "Pending For Data Validation",
    )
    for status_value in pending_user_statuses:
        for db_row in get_vendor_records_by_approver_status(status_value):
            record_id = str(db_row.get("record") or "").strip()
            if not record_id:
                continue
            db_full_record = get_vendor_record_by_record_id(record_id)
            if not _record_is_assigned_to_email(db_full_record, user_email):
                continue
            upsert_row(
                {
                    "record": record_id,
                    "company": str(db_row.get("company") or ""),
                    "category": str(db_row.get("category") or ""),
                    "approverStatus": str(db_row.get("approverStatus") or ""),
                    "name": str(db_row.get("name") or ""),
                    "type": str(db_row.get("type") or ""),
                    "createdAt": str(db_row.get("createdAt") or ""),
                }
            )

    for db_row in sql_rows:
        if not _is_created_in_sap_record(db_row):
            continue
        if not _record_is_assigned_to_email(db_row, user_email):
            continue
        summary_row = _map_vendor_summary_row(db_row)
        if not str(summary_row.get("record") or "").strip():
            continue
        upsert_row(summary_row)

    return build_rows()


@app.get("/api/vendor/update-hod-email-records")
def get_update_hod_email_records(
    user: dict[str, Any] = Depends(get_current_user),
) -> list[dict[str, Any]]:
    if not user_has_any_role(user, "user", "buyer"):
        raise HTTPException(status_code=403, detail="Only User role can access Update HOD Email records")

    all_rows = get_dashboard_records(user)
    visible_rows: list[dict[str, Any]] = []
    for row in all_rows:
        status_text = str(row.get("approverStatus") or "").strip().lower()
        if _is_pending_sap_code_creation_status(status_text):
            continue
        visible_rows.append(row)
    return visible_rows


@app.get("/api/vendor/history-records")
def get_vendor_history_records(
    user: dict[str, Any] = Depends(get_current_user),
) -> list[dict[str, Any]]:
    normalized_role = normalize_role(user.get("role"))
    if normalized_role in {"validator_user", "validator_hod"}:
        merged_rows: dict[str, dict[str, Any]] = {}
        role_sequence = (
            ("validator", "user")
            if normalized_role == "validator_user"
            else ("validator", "hod")
        )
        for role_key in role_sequence:
            scoped_user = dict(user)
            scoped_user["role"] = role_key
            for row in get_vendor_history_records(scoped_user):
                row_key = str(row.get("record") or "").strip().lower()
                if not row_key:
                    continue
                existing_row = merged_rows.get(row_key)
                if not existing_row or str(row.get("createdAt") or "") >= str(existing_row.get("createdAt") or ""):
                    merged_rows[row_key] = row
        rows = list(merged_rows.values())
        rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
        return rows

    user_email = str(user.get("email") or "").strip().lower()
    if normalized_role != "admin" and not user_email:
        return []

    merged: dict[str, dict[str, Any]] = {}
    sql_rows = get_all_vendor_records()

    def upsert_row(row: dict[str, Any]) -> None:
        key = str(row.get("record") or "").strip().lower()
        if not key:
            return
        existing = merged.get(key)
        if not existing or str(row.get("createdAt") or "") >= str(existing.get("createdAt") or ""):
            merged[key] = row

    def add_row(record: dict[str, Any] | None, default_status: str = "") -> None:
        upsert_row(_map_vendor_summary_row(record, default_status=default_status))

    if normalized_role in {"admin", "validator"}:
        for invite in invite_forms.values():
            add_row(invite, default_status="Vendor Invited")
        for form in vendor_forms.values():
            add_row(form, default_status="Pending For User Approval")
        for db_row in sql_rows:
            add_row(db_row)
    elif normalized_role == "hod":
        for invite in invite_forms.values():
            if _extract_record_hod_email(invite) != user_email:
                continue
            add_row(invite, default_status="Vendor Invited")
        for form in vendor_forms.values():
            if _extract_record_hod_email(form) != user_email:
                continue
            add_row(form, default_status="Pending For User Approval")
        for db_row in sql_rows:
            if _extract_record_hod_email(db_row) != user_email:
                continue
            add_row(db_row)
    elif normalized_role in {"user", "buyer"}:
        for invite in invite_forms.values():
            if not _record_is_assigned_to_email(invite, user_email):
                continue
            add_row(invite, default_status="Vendor Invited")
        for form in vendor_forms.values():
            if not _record_is_assigned_to_email(form, user_email):
                continue
            add_row(form, default_status="Pending For User Approval")
        for db_row in sql_rows:
            if not _record_is_assigned_to_email(db_row, user_email):
                continue
            add_row(db_row)

        if not merged:
            # Fallback for older records where assignee metadata might be incomplete.
            for invite in invite_forms.values():
                if _is_pending_user_approval_status_value(_extract_vendor_status_text(invite, default_status="Vendor Invited")):
                    add_row(invite, default_status="Vendor Invited")
            for form in vendor_forms.values():
                if _is_pending_user_approval_status_value(_extract_vendor_status_text(form, default_status="Pending For User Approval")):
                    add_row(form, default_status="Pending For User Approval")
            for db_row in sql_rows:
                if _is_pending_user_approval_status_value(_extract_vendor_status_text(db_row)):
                    add_row(db_row)
    else:
        for invite in invite_forms.values():
            if not _record_is_assigned_to_email(invite, user_email):
                continue
            add_row(invite, default_status="Vendor Invited")
        for form in vendor_forms.values():
            if not _record_is_assigned_to_email(form, user_email):
                continue
            add_row(form, default_status="Pending For User Approval")
        for db_row in sql_rows:
            if not _record_is_assigned_to_email(db_row, user_email):
                continue
            add_row(db_row)

    rows = list(merged.values())
    rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
    return rows


def _normalize_history_export_column_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _build_history_export_document_column_token_set() -> set[str]:
    tokens: set[str] = set()

    for field_name, aliases in _document_source_aliases().items():
        for candidate in (field_name, *aliases):
            candidate_text = str(candidate or "").strip()
            if not candidate_text:
                continue

            normalized = _normalize_history_export_column_key(candidate_text)
            if normalized:
                tokens.add(normalized)
            for suffix in ("FileName", "ContentType", "Data", "Downloaded"):
                suffix_normalized = _normalize_history_export_column_key(f"{candidate_text}_{suffix}")
                if suffix_normalized:
                    tokens.add(suffix_normalized)

    for column_name in IMPORT_FORM_UPLOAD_TO_COLUMN.values():
        normalized = _normalize_history_export_column_key(column_name)
        if normalized:
            tokens.add(normalized)
        for suffix in ("FileName", "ContentType", "Data", "Downloaded"):
            suffix_normalized = _normalize_history_export_column_key(f"{column_name}_{suffix}")
            if suffix_normalized:
                tokens.add(suffix_normalized)

    return tokens


HISTORY_EXPORT_DOCUMENT_COLUMN_TOKENS = _build_history_export_document_column_token_set()
HISTORY_EXPORT_DOCUMENT_NAME_HINTS = (
    "attachment",
    "document",
    "certificate",
    "certi",
    "cheque",
)


def _is_document_column_for_history_export(column_name: Any) -> bool:
    column_text = str(column_name or "").strip()
    if not column_text:
        return False

    lowered = column_text.lower()
    normalized = _normalize_history_export_column_key(column_text)
    if normalized and normalized in HISTORY_EXPORT_DOCUMENT_COLUMN_TOKENS:
        return True

    if lowered.endswith(("_filename", "filename", "_contenttype", "contenttype", "_data", "_downloaded")):
        return True
    if any(hint in lowered for hint in HISTORY_EXPORT_DOCUMENT_NAME_HINTS):
        return True

    return False


def _coerce_excel_cell_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (bytes, bytearray, memoryview)):
        return ""
    if isinstance(value, (dict, list, tuple, set)):
        try:
            return json.dumps(value, ensure_ascii=False)
        except Exception:
            return str(value)
    return value


def _extract_vendor_name_for_history_export(record: dict[str, Any] | None) -> str:
    if not isinstance(record, dict):
        return ""
    for key in ("vendorName", "Vendorname", "VendorName", "name", "Name"):
        value = str(record.get(key) or "").strip()
        if value:
            return value
    return ""


def _sanitize_download_filename_part(value: Any, fallback: str) -> str:
    text = str(value or "").strip()
    text = re.sub(r'[<>:"/\\|?*]+', "-", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text or fallback


def _build_record_export_filename_base(record_id: str, record: dict[str, Any] | None) -> str:
    safe_record = _sanitize_download_filename_part(record_id, "Record")
    vendor_name = _extract_vendor_name_for_history_export(record)
    safe_vendor = _sanitize_download_filename_part(vendor_name, "Vendor")
    return f"{safe_record} + {safe_vendor}"


def _get_history_record_export_context(
    user: dict[str, Any],
    record_id: str,
) -> tuple[str, dict[str, Any], dict[str, Any] | None, dict[str, Any] | None]:
    normalized_role = normalize_role(user.get("role"))
    user_email = str(user.get("email") or "").strip().lower()
    if not user_has_any_role(user, "admin") and not user_email:
        raise HTTPException(status_code=403, detail="Unable to identify logged-in user")

    requested_record_id = _normalize_record_id(record_id)
    if not requested_record_id:
        raise HTTPException(status_code=400, detail="recordId is required")

    summary_rows: list[dict[str, Any]] = []
    if user_email:
        summary_rows = get_vendor_records_by_assignee_email(user_email)
    summary_row_by_key: dict[str, dict[str, Any]] = {}
    for row in summary_rows:
        row_record_id = _normalize_record_id(row.get("record"))
        if not row_record_id:
            continue
        row_key = row_record_id.lower()
        if row_key not in summary_row_by_key:
            summary_row_by_key[row_key] = row

    requested_key = requested_record_id.lower()
    summary_row = summary_row_by_key.get(requested_key)

    local_record = find_local_vendor_record_by_record_id(requested_record_id)
    db_record = get_vendor_record_by_record_id(requested_record_id)
    if not local_record and not db_record and not summary_row:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    if user_has_any_role(user, "admin", "validator"):
        is_allowed = True
    elif user_has_any_role(user, "hod"):
        is_allowed = user_email in {
            email
            for email in (
                _extract_record_hod_email(local_record),
                _extract_record_hod_email(db_record),
            )
            if email
        }
    else:
        is_allowed = bool(summary_row) or _record_is_assigned_to_email(local_record, user_email) or _record_is_assigned_to_email(
            db_record, user_email
        )

    if not is_allowed:
        raise HTTPException(status_code=403, detail="Access denied for this record")

    merged_record: dict[str, Any] = {}
    if db_record:
        merged_record.update(db_record)
    if local_record:
        merged_record.update(local_record)
    if not merged_record and summary_row:
        merged_record = {
            "record": str(summary_row.get("record") or requested_record_id),
            "company": str(summary_row.get("company") or ""),
            "category": str(summary_row.get("category") or ""),
            "approverStatus": str(summary_row.get("approverStatus") or ""),
            "name": str(summary_row.get("name") or ""),
            "type": str(summary_row.get("type") or ""),
            "createdAt": str(summary_row.get("createdAt") or ""),
        }

    merged_record["recordId"] = requested_record_id
    return requested_record_id, merged_record, local_record, db_record


def _filter_record_for_history_export(record: dict[str, Any]) -> dict[str, Any]:
    filtered_record: dict[str, Any] = {}
    for key, value in (record or {}).items():
        column_name = str(key or "").strip()
        if not column_name or _is_document_column_for_history_export(column_name):
            continue
        filtered_record[column_name] = value
    return filtered_record


def _build_single_record_excel_bytes(record: dict[str, Any]) -> bytes:
    filtered_record = _filter_record_for_history_export(record)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Vendor Record"

    if not filtered_record:
        worksheet.append(["No records found"])
        worksheet.column_dimensions["A"].width = 28
    else:
        headers = list(filtered_record.keys())
        worksheet.append(headers)
        values: list[Any] = []
        column_widths = [len(str(header or "")) for header in headers]

        for idx, header in enumerate(headers):
            cell_value = _coerce_excel_cell_value(filtered_record.get(header))
            values.append(cell_value)
            text_length = len(str(cell_value or ""))
            if text_length > column_widths[idx]:
                column_widths[idx] = text_length

        worksheet.append(values)
        for idx, width in enumerate(column_widths, start=1):
            worksheet.column_dimensions[get_column_letter(idx)].width = min(80, max(12, width + 2))
        worksheet.freeze_panes = "A2"

    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _to_sap_text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (list, tuple, set)):
        parts = [str(item).strip() for item in value if str(item).strip()]
        return ", ".join(parts)
    return str(value).strip()


def _normalize_sap_tax_code(value: Any) -> str:
    raw_text = _to_sap_text_value(value)
    if not raw_text:
        return ""

    try:
        parsed = json.loads(raw_text)
    except Exception:
        parsed = None

    if isinstance(parsed, list):
        parts = [_to_sap_text_value(item) for item in parsed]
        parts = [item for item in parts if item]
        return ", ".join(parts)
    if parsed is not None and not isinstance(parsed, dict):
        return _to_sap_text_value(parsed)

    if raw_text.startswith("[") and raw_text.endswith("]"):
        inner = raw_text[1:-1]
        parts = [
            item.strip().strip("'").strip('"')
            for item in inner.split(",")
            if item.strip().strip("'").strip('"')
        ]
        return ", ".join(parts)
    return raw_text


def _build_sap_template_row_values(record: dict[str, Any]) -> dict[str, str]:
    source = record if isinstance(record, dict) else {}

    def pick(*keys: str) -> Any:
        return _first_existing_record_value(source, None, *keys)

    vendor_name = _to_sap_text_value(
        pick("vendorName", "Vendorname", "VendorName", "name", "Name")
    )
    vendor_email = _extract_record_vendor_email(source) or _to_sap_text_value(
        pick(
            "contactPersonEmail",
            "ContactPersonEmail",
            "ContactEmailID",
            "vendorEmail",
            "VendorEmail",
            "AlternativePersonEmail",
            "alternativePersonEmail",
        )
    )

    values: dict[str, Any] = {
        "Vendor_Type": pick("vendorType", "VendorType", "type", "Type"),
        "VendorName1": vendor_name,
        "VendorName2": pick("tradeNameGst", "TradeName", "tradeNamePan", "TradeNamePAN", "vendorName", "Vendorname"),
        "GST_Number": pick("gstNumber", "GSTIN"),
        "PAN_Number": pick("panNumber", "PANNo"),
        "PANName": pick("tradeNamePan", "TradeNamePAN", "panName", "PANName", "vendorName", "Vendorname"),
        "UdyamNumber": pick("udyamNumber", "UdyamNumber"),
        "MSMECategory": pick("msmeCategory", "MSMECategory"),
        "MSMEIndustry": pick("msmeIndustry", "MSMEIndustry"),
        "VendorCIN": pick("cinNumber", "CIN"),
        "Address1": pick("addressLane1", "Address1"),
        "Address2": pick("addressLane2", "Address2"),
        "Address3": pick("addressLane3", "Address3", "Adddress3"),
        "PostalCode": pick("pincode", "Pincode"),
        "City": pick("city", "City", "district", "District"),
        "Country": pick("country", "Country"),
        "BankName": pick("bankName", "BankName"),
        "Branch": pick("branchName", "BranchName"),
        "BankAccount": pick("bankAccountNumber", "BankAccount"),
        "IFSC": pick("ifscCode", "IFSC_Code"),
        "IncoTerms": pick("incoTerms", "IncoTerms", "Incoterms"),
        "PaymentTerms": pick("paymentTerms", "PaymentTerms"),
        "Mobile": pick("contactPersonMobile", "ContactPersonMobile", "alternativePersonMobile", "AlternativePersonMobile"),
        "TaxCode": _normalize_sap_tax_code(pick("TaxCode", "taxCode", "withholdingTax", "WithholdingTax", "Withholding Tax")),
        "VendorEmail": vendor_email,
        "CompanyDealing": pick("companyName", "CompanyName", "company", "Company"),
        "StateName": pick("state", "State", "StateName"),
        "TurnoverLimit": pick("turnoverLimit", "TurnoverLimit", "Turnover Limit"),
        "MSMEDate": pick("MSMEDate", "msmeDate", "udyamRegistrationDate", "UdyamRegistrationDate"),
    }

    return {header: _to_sap_text_value(values.get(header)) for header in SAP_VENDOR_TEMPLATE_HEADERS}


def _append_vendor_row_to_sap_template(record: dict[str, Any], *, clear_existing_rows: bool = False) -> int:
    if not SAP_VENDOR_TEMPLATE_PATH.exists():
        raise RuntimeError(f"SAP vendor template not found: {SAP_VENDOR_TEMPLATE_PATH}")

    workbook = load_workbook(SAP_VENDOR_TEMPLATE_PATH)
    try:
        worksheet = workbook[workbook.sheetnames[0]]
        header_to_column: dict[str, int] = {}
        for column_index in range(1, worksheet.max_column + 1):
            header_value = _to_sap_text_value(worksheet.cell(row=1, column=column_index).value)
            if header_value:
                header_to_column[header_value] = column_index

        missing_headers = [header for header in SAP_VENDOR_TEMPLATE_HEADERS if header not in header_to_column]
        if missing_headers:
            raise RuntimeError(
                "Missing SAP template headers: " + ", ".join(missing_headers)
            )

        if clear_existing_rows and worksheet.max_row > 1:
            worksheet.delete_rows(2, worksheet.max_row - 1)

        row_values = _build_sap_template_row_values(record)
        target_row = worksheet.max_row + 1
        for header in SAP_VENDOR_TEMPLATE_HEADERS:
            worksheet.cell(
                row=target_row,
                column=header_to_column[header],
                value=row_values.get(header, ""),
            )

        workbook.save(SAP_VENDOR_TEMPLATE_PATH)
        return target_row
    finally:
        workbook.close()


def _normalize_excel_2d_values(raw_values: Any) -> list[list[Any]]:
    if raw_values is None:
        return []
    if isinstance(raw_values, (tuple, list)):
        if not raw_values:
            return []
        first_item = raw_values[0]
        if isinstance(first_item, (tuple, list)):
            return [list(row) for row in raw_values]
        return [list(raw_values)]
    return [[raw_values]]


def _extract_non_empty_rows_from_openpyxl_sheet(worksheet: Any) -> list[list[Any]]:
    rows: list[list[Any]] = []
    max_row = int(getattr(worksheet, "max_row", 0) or 0)
    max_col = int(getattr(worksheet, "max_column", 0) or 0)
    if max_row <= 0 or max_col <= 0:
        return rows

    for row_idx in range(1, max_row + 1):
        row_values = [worksheet.cell(row=row_idx, column=col_idx).value for col_idx in range(1, max_col + 1)]
        if any(_to_sap_text_value(cell_value) != "" for cell_value in row_values):
            rows.append(row_values)
    return rows


def _copy_input_file_to_vendor_upload_without_refresh() -> dict[str, Any]:
    if not SAP_CONFIG_WORKBOOK_PATH.exists():
        raise RuntimeError(f"Config workbook not found: {SAP_CONFIG_WORKBOOK_PATH}")
    if not SAP_VENDOR_UPLOAD_PATH.exists():
        raise RuntimeError(f"Vendor upload workbook not found: {SAP_VENDOR_UPLOAD_PATH}")

    config_workbook = load_workbook(SAP_CONFIG_WORKBOOK_PATH, data_only=True)
    try:
        if SAP_CONFIG_INPUT_SHEET_NAME not in config_workbook.sheetnames:
            raise RuntimeError(f"Sheet '{SAP_CONFIG_INPUT_SHEET_NAME}' not found in config workbook")
        input_sheet = config_workbook[SAP_CONFIG_INPUT_SHEET_NAME]
        source_rows = _extract_non_empty_rows_from_openpyxl_sheet(input_sheet)
    finally:
        config_workbook.close()

    if not source_rows:
        raise RuntimeError(
            f"No data found in sheet '{SAP_CONFIG_INPUT_SHEET_NAME}' (fallback mode)"
        )

    max_columns = max(len(row) for row in source_rows)
    normalized_rows: list[list[Any]] = []
    for row in source_rows:
        row_values = list(row)
        if len(row_values) < max_columns:
            row_values.extend([None] * (max_columns - len(row_values)))
        normalized_rows.append(row_values)

    upload_workbook = load_workbook(SAP_VENDOR_UPLOAD_PATH)
    try:
        if SAP_CONFIG_INPUT_SHEET_NAME in upload_workbook.sheetnames:
            upload_sheet = upload_workbook[SAP_CONFIG_INPUT_SHEET_NAME]
        else:
            upload_sheet = upload_workbook[upload_workbook.sheetnames[0]]

        existing_max_row = int(getattr(upload_sheet, "max_row", 0) or 0)
        existing_max_col = int(getattr(upload_sheet, "max_column", 0) or 0)
        if existing_max_row > 0 and existing_max_col > 0:
            for row_idx in range(1, existing_max_row + 1):
                for col_idx in range(1, existing_max_col + 1):
                    upload_sheet.cell(row=row_idx, column=col_idx, value=None)

        for row_idx, row_values in enumerate(normalized_rows, start=1):
            for col_idx, cell_value in enumerate(row_values, start=1):
                upload_sheet.cell(row=row_idx, column=col_idx, value=cell_value)

        upload_workbook.save(SAP_VENDOR_UPLOAD_PATH)
        return {
            "rowsCopied": len(normalized_rows),
            "columnsCopied": max_columns,
            "sourceSheet": SAP_CONFIG_INPUT_SHEET_NAME,
            "targetSheet": upload_sheet.title,
            "refreshMode": "fallback_no_com",
        }
    finally:
        upload_workbook.close()


def _refresh_power_query_and_sync_vendor_upload_via_powershell() -> dict[str, Any]:
    script_text = r"""
param(
  [Parameter(Mandatory=$true)][string]$ConfigPath,
  [Parameter(Mandatory=$true)][string]$UploadPath,
  [Parameter(Mandatory=$true)][string]$SheetName,
  [Parameter(Mandatory=$true)][int]$TimeoutSec
)

$ErrorActionPreference = 'Stop'
$excel = $null
$configBook = $null
$uploadBook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false

  $configBook = $excel.Workbooks.Open($ConfigPath, 0, $false)
  $connCount = 0
  try { $connCount = [int]$configBook.Connections.Count } catch {}
  for ($i = 1; $i -le $connCount; $i++) {
    try {
      $conn = $configBook.Connections.Item($i)
      try { $conn.OLEDBConnection.BackgroundQuery = $false } catch {}
      try { $conn.ODBCConnection.BackgroundQuery = $false } catch {}
    } catch {}
  }

  $configBook.RefreshAll()
  try { $excel.CalculateUntilAsyncQueriesDone() } catch {}

  $deadline = (Get-Date).AddSeconds([Math]::Max(5, $TimeoutSec))
  while ($true) {
    $refreshing = $false
    for ($i = 1; $i -le $connCount; $i++) {
      try {
        $conn = $configBook.Connections.Item($i)
        try { if ($conn.OLEDBConnection.Refreshing) { $refreshing = $true; break } } catch {}
        try { if ($conn.ODBCConnection.Refreshing) { $refreshing = $true; break } } catch {}
      } catch {}
    }

    $calcDone = $true
    try { $calcDone = ([int]$excel.CalculationState -eq 0) } catch {}

    if (-not $refreshing -and $calcDone) { break }
    if ((Get-Date) -ge $deadline) {
      throw "Power Query refresh timed out after $TimeoutSec seconds"
    }
    Start-Sleep -Seconds 1
  }

  $inputSheet = $configBook.Worksheets.Item($SheetName)
  $region = $inputSheet.Range("A1").CurrentRegion
  $rows = [int]$region.Rows.Count
  $cols = [int]$region.Columns.Count
  if ($rows -lt 1 -or $cols -lt 1) {
    throw "No data found in sheet '$SheetName' after query refresh"
  }
  $values = $region.Value2
  if ($null -eq $values) {
    throw "No data found in sheet '$SheetName' after query refresh"
  }

  $uploadBook = $excel.Workbooks.Open($UploadPath, 0, $false)
  try { $targetSheet = $uploadBook.Worksheets.Item($SheetName) } catch { $targetSheet = $uploadBook.Worksheets.Item(1) }

  $targetSheet.Cells.ClearContents() | Out-Null
  $dest = $targetSheet.Range($targetSheet.Cells.Item(1, 1), $targetSheet.Cells.Item($rows, $cols))
  $dest.Value2 = $values
  $uploadBook.Save()

  [PSCustomObject]@{
    rowsCopied = $rows
    columnsCopied = $cols
    sourceSheet = $SheetName
    targetSheet = [string]$targetSheet.Name
    refreshMode = "com_refresh_powershell"
  } | ConvertTo-Json -Compress
}
catch {
  Write-Error $_.Exception.Message
  exit 1
}
finally {
  if ($uploadBook -ne $null) { try { $uploadBook.Close($false) } catch {} }
  if ($configBook -ne $null) { try { $configBook.Close($false) } catch {} }
  if ($excel -ne $null) { try { $excel.Quit() } catch {} }
}
"""

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as temp_file:
            temp_file.write(script_text)
            temp_path = temp_file.name

        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                temp_path,
                "-ConfigPath",
                str(SAP_CONFIG_WORKBOOK_PATH),
                "-UploadPath",
                str(SAP_VENDOR_UPLOAD_PATH),
                "-SheetName",
                SAP_CONFIG_INPUT_SHEET_NAME,
                "-TimeoutSec",
                str(SAP_QUERY_REFRESH_TIMEOUT_SECONDS),
            ],
            capture_output=True,
            text=True,
            timeout=SAP_QUERY_REFRESH_TIMEOUT_SECONDS + 180,
        )
    finally:
        if temp_path:
            try:
                os.remove(temp_path)
            except Exception:
                pass

    if completed.returncode != 0:
        stdout_text = (completed.stdout or "").strip()
        stderr_text = (completed.stderr or "").strip()
        details = stderr_text or stdout_text or "Unknown PowerShell COM refresh error"
        raise RuntimeError(details)

    stdout_lines = [line.strip() for line in (completed.stdout or "").splitlines() if line.strip()]
    if not stdout_lines:
        raise RuntimeError("PowerShell COM refresh did not return any output")

    last_line = stdout_lines[-1]
    try:
        payload = json.loads(last_line)
    except Exception as exc:
        raise RuntimeError(f"Invalid PowerShell COM refresh response: {last_line}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("Invalid PowerShell COM refresh response payload")
    return payload


def _excel_connection_is_refreshing(connection: Any) -> bool:
    for attribute_name in ("OLEDBConnection", "ODBCConnection"):
        try:
            child_connection = getattr(connection, attribute_name)
            if bool(getattr(child_connection, "Refreshing", False)):
                return True
        except Exception:
            continue
    return False


def _set_excel_connections_synchronous(workbook: Any) -> None:
    try:
        connection_count = int(getattr(workbook.Connections, "Count", 0))
    except Exception:
        connection_count = 0
    for index in range(1, connection_count + 1):
        try:
            connection = workbook.Connections.Item(index)
        except Exception:
            continue
        for attribute_name in ("OLEDBConnection", "ODBCConnection"):
            try:
                child_connection = getattr(connection, attribute_name)
                setattr(child_connection, "BackgroundQuery", False)
            except Exception:
                continue


def _wait_for_excel_refresh_completion(
    excel_app: Any,
    workbook: Any,
    timeout_seconds: int = SAP_QUERY_REFRESH_TIMEOUT_SECONDS,
) -> None:
    timeout_value = max(5, int(timeout_seconds))
    deadline = time.time() + timeout_value

    while True:
        refreshing = False
        try:
            connection_count = int(getattr(workbook.Connections, "Count", 0))
        except Exception:
            connection_count = 0
        for index in range(1, connection_count + 1):
            try:
                connection = workbook.Connections.Item(index)
            except Exception:
                continue
            if _excel_connection_is_refreshing(connection):
                refreshing = True
                break

        try:
            calculation_done = int(getattr(excel_app, "CalculationState", 0)) == 0
        except Exception:
            calculation_done = True

        if not refreshing and calculation_done:
            return

        if time.time() >= deadline:
            raise RuntimeError(
                f"Power Query refresh timed out after {timeout_value} seconds"
            )
        time.sleep(1)


def _refresh_power_query_and_sync_vendor_upload() -> dict[str, Any]:
    if not SAP_CONFIG_WORKBOOK_PATH.exists():
        raise RuntimeError(f"Config workbook not found: {SAP_CONFIG_WORKBOOK_PATH}")
    if not SAP_VENDOR_UPLOAD_PATH.exists():
        raise RuntimeError(f"Vendor upload workbook not found: {SAP_VENDOR_UPLOAD_PATH}")

    pythoncom = None
    win32com = None
    try:
        import pythoncom as _pythoncom  # type: ignore[import-untyped]
        import win32com.client as _win32com_client  # type: ignore[import-untyped]
        pythoncom = _pythoncom
        win32com = _win32com_client
    except Exception as exc:
        try:
            return _refresh_power_query_and_sync_vendor_upload_via_powershell()
        except Exception as ps_exc:
            fallback_result = _copy_input_file_to_vendor_upload_without_refresh()
            fallback_result["refreshWarning"] = (
                "Power Query refresh skipped because pywin32 import failed and PowerShell COM refresh was unavailable"
            )
            fallback_result["comError"] = f"pywin32 import failed: {exc}; PowerShell COM error: {ps_exc}"
            fallback_result["pythonExecutable"] = os.sys.executable
            return fallback_result

    def fallback_after_com_error(primary_error: Exception) -> dict[str, Any]:
        try:
            return _refresh_power_query_and_sync_vendor_upload_via_powershell()
        except Exception as ps_exc:
            fallback_result = _copy_input_file_to_vendor_upload_without_refresh()
            fallback_result["refreshWarning"] = (
                "Power Query refresh skipped because Excel COM refresh was unavailable in current session"
            )
            fallback_result["comError"] = f"pywin32 COM error: {primary_error}; PowerShell COM error: {ps_exc}"
            fallback_result["pythonExecutable"] = os.sys.executable
            return fallback_result

    pythoncom.CoInitialize()
    try:
        excel_app = None
        config_workbook = None
        upload_workbook = None
        com_error: Exception | None = None
        try:
            excel_app = win32com.DispatchEx("Excel.Application")
            excel_app.Visible = True
            excel_app.DisplayAlerts = False
            excel_app.AskToUpdateLinks = False
            excel_app.EnableEvents = False

            config_workbook = excel_app.Workbooks.Open(
                str(SAP_CONFIG_WORKBOOK_PATH),
                UpdateLinks=0,
                ReadOnly=False,
            )
            _set_excel_connections_synchronous(config_workbook)
            config_workbook.RefreshAll()
            try:
                excel_app.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            _wait_for_excel_refresh_completion(excel_app, config_workbook)

            input_sheet = config_workbook.Worksheets(SAP_CONFIG_INPUT_SHEET_NAME)
            input_region = input_sheet.Range("A1").CurrentRegion
            source_rows = _normalize_excel_2d_values(input_region.Value)
            source_rows = [
                row
                for row in source_rows
                if any(_to_sap_text_value(cell) != "" for cell in row)
            ]
            if not source_rows:
                raise RuntimeError(
                    f"No data found in sheet '{SAP_CONFIG_INPUT_SHEET_NAME}' after query refresh"
                )

            max_columns = max(len(row) for row in source_rows)
            normalized_rows: list[list[Any]] = []
            for row in source_rows:
                row_values = list(row)
                if len(row_values) < max_columns:
                    row_values.extend([None] * (max_columns - len(row_values)))
                normalized_rows.append(row_values)

            upload_workbook = excel_app.Workbooks.Open(
                str(SAP_VENDOR_UPLOAD_PATH),
                UpdateLinks=0,
                ReadOnly=False,
            )
            try:
                upload_sheet = upload_workbook.Worksheets(SAP_CONFIG_INPUT_SHEET_NAME)
            except Exception:
                upload_sheet = upload_workbook.Worksheets(1)

            upload_sheet.Cells.ClearContents()
            row_count = len(normalized_rows)
            target_range = upload_sheet.Range(
                upload_sheet.Cells(1, 1),
                upload_sheet.Cells(row_count, max_columns),
            )
            target_range.Value = tuple(tuple(row) for row in normalized_rows)
            upload_workbook.Save()
            return {
                "rowsCopied": row_count,
                "columnsCopied": max_columns,
                "sourceSheet": SAP_CONFIG_INPUT_SHEET_NAME,
                "targetSheet": str(getattr(upload_sheet, "Name", "")),
                "refreshMode": "com_refresh",
            }
        except Exception as exc:
            com_error = exc
        finally:
            if upload_workbook is not None:
                try:
                    upload_workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if config_workbook is not None:
                try:
                    config_workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel_app is not None:
                try:
                    excel_app.Quit()
                except Exception:
                    pass

        if com_error is not None:
            return fallback_after_com_error(com_error)
        raise RuntimeError("Unknown Excel COM refresh error")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _build_sap_source_record_by_record_id(record_id: str) -> dict[str, Any]:
    source_record: dict[str, Any] = {}
    db_record = get_vendor_record_by_record_id(record_id)
    local_record = find_local_vendor_record_by_record_id(record_id)
    if db_record:
        source_record.update(db_record)
    if local_record:
        source_record.update(local_record)
    if source_record and not _normalize_record_id(source_record.get("recordId")):
        source_record["recordId"] = record_id
    return source_record


def _queue_sap_sync_for_pending_status(
    background_tasks: BackgroundTasks | None,
    record_id: str,
    status_value: Any,
    *,
    force: bool = False,
) -> tuple[bool, int]:
    if background_tasks is None:
        return False, 0
    if not force and not _is_pending_sap_code_creation_status(status_value):
        return False, 0

    normalized_record_id = _normalize_record_id(record_id)
    if not normalized_record_id:
        return False, 0

    background_tasks.add_task(_run_sap_sync_pipeline_for_records_background, [normalized_record_id])
    return True, 1


def _persist_sap_vendor_code_for_record(record_id: str, vendor_code: Any) -> dict[str, Any]:
    normalized_record_id = _normalize_record_id(record_id)
    code_text = str(vendor_code or "").strip()
    if not normalized_record_id or not code_text:
        raise RuntimeError("recordId and vendorCode are required")

    try:
        ensure_vendor_columns(["VendorCode", "SAPCodeGeneratedAt"])
    except Exception:
        pass

    generated_at = now_iso()
    updates = {
        "vendorCode": code_text,
        "VendorCode": code_text,
        "SAPVendorCode": code_text,
        "sapVendorCode": code_text,
        "BusinessPartnerCode": code_text,
        "businessPartnerCode": code_text,
        "SAPCodeGeneratedAt": generated_at,
        "sapCodeGeneratedAt": generated_at,
    }

    db_record = get_vendor_record_by_record_id(normalized_record_id)
    db_updated = update_vendor_record_by_record_id(normalized_record_id, updates)
    if db_record and not db_updated:
        raise RuntimeError("Vendor record found in database but SAP vendor code update failed")

    local_updated = update_local_vendor_record_by_record_id(normalized_record_id, updates)
    if not db_updated and not local_updated:
        raise RuntimeError("Unable to update SAP vendor code for vendor record")

    return {
        "recordId": normalized_record_id,
        "vendorCode": code_text,
        "dbUpdated": db_updated,
        "localUpdated": local_updated,
    }


def _send_vendor_code_created_email_for_record(record_id: str, vendor_code: Any) -> dict[str, Any]:
    normalized_record_id = _normalize_record_id(record_id)
    code_text = str(vendor_code or "").strip()
    if not normalized_record_id or not code_text:
        raise RuntimeError("recordId and vendorCode are required for vendor code notification")

    local_record = find_local_vendor_record_by_record_id(normalized_record_id)
    db_record = get_vendor_record_by_record_id(normalized_record_id)
    recipient_name, recipient_email, vendor_name = _resolve_vendor_code_notification_details(local_record, db_record)
    if not recipient_email:
        raise RuntimeError("Unable to determine recipient email for vendor code notification")

    send_vendor_code_created_email(
        recipient_name=recipient_name,
        recipient_email=recipient_email,
        vendor_name=vendor_name,
        vendor_code=code_text,
        record_id=normalized_record_id,
    )
    return {
        "recordId": normalized_record_id,
        "recipientName": recipient_name,
        "recipientEmail": recipient_email,
        "vendorName": vendor_name,
        "vendorCode": code_text,
    }


def _run_sap_sync_pipeline_for_records(record_ids: list[str]) -> dict[str, Any]:
    normalized_record_ids: list[str] = []
    seen_record_ids: set[str] = set()
    for raw_record_id in record_ids:
        record_id = _normalize_record_id(raw_record_id)
        if not record_id:
            continue
        lowered = record_id.lower()
        if lowered in seen_record_ids:
            continue
        seen_record_ids.add(lowered)
        normalized_record_ids.append(record_id)

    if not normalized_record_ids:
        return {
            "queuedCount": 0,
            "templateRowsWritten": 0,
            "templateErrors": [],
            "vendorUploadSync": None,
            "vendorUploadError": None,
            "vendorUploadOpened": False,
            "vendorUploadOpenPath": str(SAP_VENDOR_UPLOAD_PATH),
            "vendorUploadOpenError": None,
            "vendorUploadMinimized": False,
            "vendorUploadMinimizePath": str(SAP_VENDOR_UPLOAD_PATH),
            "vendorUploadMinimizeError": None,
            "vendorUploadClosed": False,
            "vendorUploadClosePath": str(SAP_VENDOR_UPLOAD_PATH),
            "vendorUploadCloseError": None,
            "sapLogonStarted": False,
            "sapLogonPath": str(SAP_LOGON_EXECUTABLE_PATH),
            "sapLogonError": None,
            "sapServerName": SAP_SERVER_NAME,
            "sapUser": SAP_LOGON_USERNAME,
            "sapClient": SAP_LOGON_CLIENT,
            "sapAutoLogonSuccess": False,
            "sapAutoLogonError": None,
            "sapSecurityPopupsConfirmed": 0,
            "sapTransactionCode": SAP_VENDOR_UPLOAD_TRANSACTION_CODE,
            "sapExcelMinimizeDelaySeconds": SAP_EXCEL_MINIMIZE_DELAY_SECONDS,
            "sapTransactionSubmitted": False,
            "sapTransactionError": None,
            "sapVendorCode": None,
            "sapVendorCodePopupFound": False,
            "sapVendorCodePopupTitle": None,
            "sapVendorCodePopupText": None,
            "sapVendorCodePopupConfirmed": False,
            "sapVendorCodeUpdateResults": [],
            "sapVendorCodeUpdateErrors": [],
            "sapVendorCodeEmailResults": [],
            "sapVendorCodeEmailErrors": [],
            "sapClosed": False,
            "sapCloseError": None,
        }

    with SAP_SYNC_LOCK:
        template_rows_written = 0
        template_errors: list[str] = []
        sap_vendor_code_update_results: list[dict[str, Any]] = []
        sap_vendor_code_update_errors: list[str] = []
        sap_vendor_code_email_results: list[dict[str, Any]] = []
        sap_vendor_code_email_errors: list[str] = []
        should_clear_existing_rows = True

        for record_id in normalized_record_ids:
            source_record = _build_sap_source_record_by_record_id(record_id)
            if not source_record:
                template_errors.append(f"{record_id}: Vendor record not found for SAP sync")
                continue
            try:
                _append_vendor_row_to_sap_template(
                    source_record,
                    clear_existing_rows=should_clear_existing_rows,
                )
                should_clear_existing_rows = False
                template_rows_written += 1
            except Exception as exc:
                template_errors.append(f"{record_id}: {exc}")

        vendor_upload_sync: dict[str, Any] | None = None
        vendor_upload_error: str | None = None
        post_sync_result: dict[str, Any] = {
            "vendorUploadOpened": False,
            "vendorUploadOpenPath": str(SAP_VENDOR_UPLOAD_PATH),
            "vendorUploadOpenError": None,
            "vendorUploadMinimized": False,
            "vendorUploadMinimizePath": str(SAP_VENDOR_UPLOAD_PATH),
            "vendorUploadMinimizeError": None,
            "vendorUploadClosed": False,
            "vendorUploadClosePath": str(SAP_VENDOR_UPLOAD_PATH),
            "vendorUploadCloseError": None,
            "sapLogonStarted": False,
            "sapLogonPath": str(SAP_LOGON_EXECUTABLE_PATH),
            "sapLogonError": None,
            "sapServerName": SAP_SERVER_NAME,
            "sapUser": SAP_LOGON_USERNAME,
            "sapClient": SAP_LOGON_CLIENT,
            "sapAutoLogonSuccess": False,
            "sapAutoLogonError": None,
            "sapSecurityPopupsConfirmed": 0,
            "sapTransactionCode": SAP_VENDOR_UPLOAD_TRANSACTION_CODE,
            "sapExcelMinimizeDelaySeconds": SAP_EXCEL_MINIMIZE_DELAY_SECONDS,
            "sapTransactionSubmitted": False,
            "sapTransactionError": None,
            "sapVendorCode": None,
            "sapVendorCodePopupFound": False,
            "sapVendorCodePopupTitle": None,
            "sapVendorCodePopupText": None,
            "sapVendorCodePopupConfirmed": False,
        }
        workbook_cleanup_result: dict[str, Any] = {
            "path": str(SAP_VENDOR_UPLOAD_PATH),
            "killed": False,
            "alreadyClosed": False,
            "error": None,
        }
        sap_cleanup_result: dict[str, Any] = {
            "killed": False,
            "alreadyClosed": False,
            "error": None,
        }
        if template_rows_written > 0:
            try:
                vendor_upload_sync = _refresh_power_query_and_sync_vendor_upload()
            except Exception as exc:
                vendor_upload_error = str(exc)
            if vendor_upload_sync is not None and not vendor_upload_error:
                try:
                    post_sync_result = run_post_vendor_upload_process(
                        vendor_upload_path=SAP_VENDOR_UPLOAD_PATH,
                        sap_logon_path=SAP_LOGON_EXECUTABLE_PATH,
                        open_vendor_upload_after_sync=SAP_OPEN_VENDOR_UPLOAD_AFTER_SYNC,
                        close_vendor_upload_after_sync=SAP_CLOSE_VENDOR_UPLOAD_AFTER_SYNC,
                        auto_launch_logon=SAP_AUTO_LAUNCH_LOGON,
                        auto_logon_after_launch=SAP_AUTO_LOGON_AFTER_LAUNCH,
                        sap_server_name=SAP_SERVER_NAME,
                        sap_username=SAP_LOGON_USERNAME,
                        sap_password=SAP_LOGON_PASSWORD,
                        sap_client=SAP_LOGON_CLIENT,
                        sap_language=SAP_LOGON_LANGUAGE,
                        sap_wait_timeout_seconds=SAP_GUI_WAIT_TIMEOUT_SECONDS,
                        sap_transaction_wait_seconds=0,
                        sap_excel_minimize_delay_seconds=SAP_EXCEL_MINIMIZE_DELAY_SECONDS,
                        sap_transaction_code=SAP_VENDOR_UPLOAD_TRANSACTION_CODE,
                    )
                    sap_vendor_code = str(post_sync_result.get("sapVendorCode") or "").strip()
                    if sap_vendor_code:
                        if len(normalized_record_ids) == 1:
                            try:
                                sap_vendor_code_update_results.append(
                                    _persist_sap_vendor_code_for_record(normalized_record_ids[0], sap_vendor_code)
                                )
                                try:
                                    sap_vendor_code_email_results.append(
                                        _send_vendor_code_created_email_for_record(
                                            normalized_record_ids[0],
                                            sap_vendor_code,
                                        )
                                    )
                                except Exception as exc:
                                    sap_vendor_code_email_errors.append(f"{normalized_record_ids[0]}: {exc}")
                            except Exception as exc:
                                sap_vendor_code_update_errors.append(f"{normalized_record_ids[0]}: {exc}")
                        else:
                            sap_vendor_code_update_errors.append(
                                "SAP returned one generated vendor code but multiple record IDs were processed"
                            )
                finally:
                    workbook_cleanup_result = kill_excel_processes()
                    sap_cleanup_result = kill_sap_processes()

    return {
        "queuedCount": len(normalized_record_ids),
        "templateRowsWritten": template_rows_written,
        "templateErrors": template_errors,
        "vendorUploadSync": vendor_upload_sync,
        "vendorUploadError": vendor_upload_error,
        "vendorUploadOpened": bool(post_sync_result.get("vendorUploadOpened")),
        "vendorUploadOpenPath": str(post_sync_result.get("vendorUploadOpenPath") or str(SAP_VENDOR_UPLOAD_PATH)),
        "vendorUploadOpenError": post_sync_result.get("vendorUploadOpenError"),
        "vendorUploadMinimized": bool(post_sync_result.get("vendorUploadMinimized")),
        "vendorUploadMinimizePath": str(post_sync_result.get("vendorUploadMinimizePath") or str(SAP_VENDOR_UPLOAD_PATH)),
        "vendorUploadMinimizeError": post_sync_result.get("vendorUploadMinimizeError"),
        "vendorUploadClosed": bool(workbook_cleanup_result.get("killed") or workbook_cleanup_result.get("alreadyClosed")),
        "vendorUploadClosePath": str(workbook_cleanup_result.get("path") or str(SAP_VENDOR_UPLOAD_PATH)),
        "vendorUploadCloseError": workbook_cleanup_result.get("error"),
        "sapLogonStarted": bool(post_sync_result.get("sapLogonStarted")),
        "sapLogonPath": str(post_sync_result.get("sapLogonPath") or str(SAP_LOGON_EXECUTABLE_PATH)),
        "sapLogonError": post_sync_result.get("sapLogonError"),
        "sapServerName": str(post_sync_result.get("sapServerName") or SAP_SERVER_NAME),
        "sapUser": str(post_sync_result.get("sapUser") or SAP_LOGON_USERNAME),
        "sapClient": str(post_sync_result.get("sapClient") or SAP_LOGON_CLIENT),
        "sapAutoLogonSuccess": bool(post_sync_result.get("sapAutoLogonSuccess")),
        "sapAutoLogonError": post_sync_result.get("sapAutoLogonError"),
        "sapSecurityPopupsConfirmed": int(post_sync_result.get("sapSecurityPopupsConfirmed") or 0),
        "sapTransactionCode": str(post_sync_result.get("sapTransactionCode") or SAP_VENDOR_UPLOAD_TRANSACTION_CODE),
        "sapExcelMinimizeDelaySeconds": int(
            post_sync_result.get("sapExcelMinimizeDelaySeconds") or SAP_EXCEL_MINIMIZE_DELAY_SECONDS
        ),
        "sapTransactionSubmitted": bool(post_sync_result.get("sapTransactionSubmitted")),
        "sapTransactionError": post_sync_result.get("sapTransactionError"),
        "sapVendorCode": post_sync_result.get("sapVendorCode"),
        "sapVendorCodePopupFound": bool(post_sync_result.get("sapVendorCodePopupFound")),
        "sapVendorCodePopupTitle": post_sync_result.get("sapVendorCodePopupTitle"),
        "sapVendorCodePopupText": post_sync_result.get("sapVendorCodePopupText"),
        "sapVendorCodePopupConfirmed": bool(post_sync_result.get("sapVendorCodePopupConfirmed")),
        "sapVendorCodeUpdateResults": sap_vendor_code_update_results,
        "sapVendorCodeUpdateErrors": sap_vendor_code_update_errors,
        "sapVendorCodeEmailResults": sap_vendor_code_email_results,
        "sapVendorCodeEmailErrors": sap_vendor_code_email_errors,
        "sapClosed": bool(sap_cleanup_result.get("killed") or sap_cleanup_result.get("alreadyClosed")),
        "sapCloseError": sap_cleanup_result.get("error"),
    }


def _run_sap_sync_pipeline_for_records_background(record_ids: list[str]) -> None:
    try:
        result = _run_sap_sync_pipeline_for_records(record_ids)
        print(
            "[SAP SYNC] completed",
            json.dumps(result, ensure_ascii=False),
        )
    except Exception as exc:
        print(f"[SAP SYNC] failed: {exc}")


def _run_sap_sync_pipeline_record_by_record_background(record_ids: list[str]) -> None:
    normalized_record_ids: list[str] = []
    seen_record_ids: set[str] = set()
    for raw_record_id in record_ids:
        record_id = _normalize_record_id(raw_record_id)
        if not record_id:
            continue
        lowered = record_id.lower()
        if lowered in seen_record_ids:
            continue
        seen_record_ids.add(lowered)
        normalized_record_ids.append(record_id)

    for record_id in normalized_record_ids:
        _run_sap_sync_pipeline_for_records_background([record_id])


PDF_VENDOR_FORM_SECTIONS: list[dict[str, Any]] = [
    {
        "title": "1. Basic Details",
        "cards": [
            {
                "title": "Basic Vendor Details",
                "fields": [
                    {"label": "Company Name", "keys": ["companyName", "CompanyName", "company", "Company"]},
                    {"label": "Vendor Name", "keys": ["vendorName", "Vendorname", "VendorName", "name", "Name"]},
                    {"label": "Vendor Category", "keys": ["vendorCategory", "VendorCategory", "Vendor_Category", "category", "Category"]},
                ],
            },
            {
                "title": "GST Details",
                "fields": [
                    {"label": "GST Registration Status", "keys": ["gstRegistrationStatus", "GSTStatus"]},
                    {"label": "GST Number", "keys": ["gstNumber", "GSTIN"]},
                    {"label": "Trade Name as per GST", "keys": ["tradeNameGst", "TradeName"]},
                    {"label": "Taxpayer Type", "keys": ["TaxpayerTypeGST", "taxpayerTypeGst", "Taxpayer Type"]},
                    {"label": "GST Dealer Type", "keys": ["gstDealerType", "GstDealerType"]},
                    {"label": "GST Return Filing", "keys": ["gstReturnFiling", "GstReturnFiling"]},
                    {"label": "Invoice Applicable", "keys": ["invoiceApplicable", "InvoiceApplicable"]},
                    {"label": "E-Invoice Applicability", "keys": ["eInvoiceApplicability", "EInvoiceApplicability"]},
                    {"label": "E-Way Bill Applicability", "keys": ["eWayBillApplicability", "EWayBillApplicability"]},
                    {"label": "Upload GST Certificate", "keys": ["gstDocument", "GSTAttachment", "GSTCertificateName"], "doc": True, "full": True},
                ],
            },
            {
                "title": "Address Details",
                "fields": [
                    {"label": "Address Lane 1", "keys": ["addressLane1", "Address1"]},
                    {"label": "Address Lane 2", "keys": ["addressLane2", "Address2"]},
                    {"label": "Address Lane 3", "keys": ["addressLane3", "Address3", "Adddress3"]},
                    {"label": "Pincode", "keys": ["pincode", "Pincode"]},
                    {"label": "Country", "keys": ["country", "Country"]},
                    {"label": "State", "keys": ["state", "State", "StateName"]},
                    {"label": "District", "keys": ["district", "District"]},
                ],
            },
            {
                "title": "Company Status & CIN Details",
                "fields": [
                    {"label": "CIN Status", "keys": ["company", "CompanyStatus", "CINStatus", "cinStatus"]},
                    {"label": "CIN Number", "keys": ["cinNumber", "CIN"]},
                    {"label": "Upload CIN Certificate", "keys": ["cinDocument", "CINCerti", "Incorporation_Filename"], "doc": True, "full": True},
                ],
            },
            {
                "title": "PAN Details",
                "fields": [
                    {"label": "PAN Status", "keys": ["panStatus", "PANStatus"]},
                    {"label": "PAN Number", "keys": ["panNumber", "PANNo"]},
                    {"label": "Business Entity Type", "keys": ["businessEntityType", "BusinessEntityType"]},
                    {"label": "Trade Name as per PAN", "keys": ["tradeNamePan", "TradeNamePAN"]},
                    {"label": "ITR Filing Status", "keys": ["itrFilingStatus", "ItrFilingStatus"]},
                    {"label": "Upload PAN Card", "keys": ["panDocument", "PANAttachment", "PAN_Card_Filename"], "doc": True, "full": True},
                ],
            },
            {
                "title": "MSME Details",
                "fields": [
                    {"label": "MSME Status", "keys": ["msmeStatus", "MSMEStatus"]},
                    {"label": "Udyam Number", "keys": ["udyamNumber", "UdyamNumber"]},
                    {"label": "MSME Category", "keys": ["msmeCategory", "MSMECategory"]},
                    {"label": "MSME Industry", "keys": ["msmeIndustry", "MSMEIndustry"]},
                    {"label": "Upload MSME Certificate", "keys": ["msmeDocument", "MSMECerti", "MSME_Certificate_Filename"], "doc": True, "full": True},
                ],
            },
            {
                "title": "PF & ESI Details",
                "fields": [
                    {"label": "PF Registration Status", "keys": ["pfStatus", "PFRegistrationStatus"]},
                    {"label": "PF Number", "keys": ["pfNumber", "PFNumber"]},
                    {"label": "Upload PF Certificate", "keys": ["pfDocument", "PFCerti", "PF_Certificate_Filename"], "doc": True, "full": True},
                    {"label": "ESI Status", "keys": ["esiStatus", "ESIStatus"]},
                    {"label": "ESI Number", "keys": ["esiNumber", "ESIN"]},
                    {"label": "Upload ESI Certificate", "keys": ["esiDocument", "ESICerti", "ESI_Filename"], "doc": True, "full": True},
                ],
            },
        ],
    },
    {
        "title": "2. Bank Details",
        "cards": [
            {
                "title": "Bank Details",
                "fields": [
                    {"label": "Payment Method RTGS", "keys": ["bankPaymentMethod", "BankPaymentMethod"]},
                    {"label": "IFSC Code", "keys": ["ifscCode", "IFSC_Code"]},
                    {"label": "Bank Name", "keys": ["bankName", "BankName"]},
                    {"label": "Branch Name", "keys": ["branchName", "BranchName"]},
                    {"label": "Bank Account Number", "keys": ["bankAccountNumber", "BankAccount"]},
                    {"label": "Cancelled Cheque / Bank Certificate", "keys": ["cancelledCheque", "CancelledCheque", "Cancelled_Cheque_Filename"], "doc": True, "full": True},
                ],
            },
        ],
    },
    {
        "title": "3. Statutory Details",
        "cards": [
            {
                "title": "Statutory Details",
                "fields": [
                    {"label": "Statutory Details", "keys": ["statutoryDetails"], "multiline": True, "full": True},
                    {"label": "Total Employees", "keys": ["totalEmployees", "TotalEmployees"]},
                    {"label": "Inhouse Machineries", "keys": ["inhouseMachineries", "InhouseMachine"]},
                    {"label": "Inhouse Machineries Document", "keys": ["inhouseMachineriesDoc", "InhouseMachineAttchment", "Machineries_List_Filename"], "doc": True, "full": True},
                    {"label": "Number Of Plants", "keys": ["numberOfPlants", "NumberofPlant"]},
                    {"label": "Number Of Plants Document", "keys": ["numberOfPlantsDoc", "TopCustomerAttachment"], "doc": True, "full": True},
                    {"label": "Year of Incorporation", "keys": ["yearOfIncorporation", "Incoporationdate"]},
                    {"label": "Selected Previous Year", "keys": ["selectedPreviousYear", "Previous_Year"]},
                    {"label": "Previous Year Turnover", "keys": ["previousYearTurnover", "PY_Turnover"]},
                    {"label": "Head Office Location", "keys": ["headOfficeLocation", "CorporateLocation"]},
                ],
            },
        ],
    },
    {
        "title": "4. Product/Service and Quality",
        "cards": [
            {
                "title": "Product/Service and Quality",
                "fields": [
                    {"label": "Material Dealings In", "keys": ["materialDealingsIn", "MaterialDealing"], "multiline": True, "full": True},
                    {"label": "Service Provide For", "keys": ["serviceProvideFor", "servicedealing"], "multiline": True, "full": True},
                    {"label": "ISO Certificates", "keys": ["isoCertificates", "ISO_Certificate"]},
                    {"label": "ISO 1", "keys": ["iso1", "ISO_File1_Name"], "doc": True},
                    {"label": "ISO 2", "keys": ["iso2", "ISO_File2_Name"], "doc": True},
                    {"label": "ISO 3", "keys": ["iso3", "ISO_File3_Name"], "doc": True},
                    {"label": "ISO 4", "keys": ["iso4", "ISO_File4_Name"], "doc": True},
                    {"label": "ISO 5", "keys": ["iso5", "ISO_File5_Name"], "doc": True},
                    {"label": "Eco Vadis Score", "keys": ["ecoVadisScore", "EcoVadis"]},
                    {"label": "Eco Vadis Document", "keys": ["ecoVadisDoc", "ECOVadis_Filename"], "doc": True, "full": True},
                    {"label": "OHSAS Certificate", "keys": ["ohsasCertificate", "OHSAS_Code"]},
                    {"label": "OHSAS Document", "keys": ["ohsasDoc", "OHSAS_Filename"], "doc": True, "full": True},
                    {"label": "Relevant Certificate", "keys": ["relevantCertificate", "OtherCerti"]},
                    {"label": "Relevant Certificate Document", "keys": ["relevantCertificateDoc", "Other_Filename"], "doc": True, "full": True},
                ],
            },
        ],
    },
    {
        "title": "5. Contact Person Details",
        "cards": [
            {
                "title": "Contact Person Details",
                "fields": [
                    {"label": "Information Furnished By", "keys": ["contactPersonName", "ContactPerson"]},
                    {"label": "Designation", "keys": ["contactPersonDesignation"]},
                    {"label": "Person Email", "keys": ["contactPersonEmail", "ContactEmailID"]},
                    {"label": "Person Mobile No.", "keys": ["contactPersonMobile"]},
                    {"label": "Alternative Person Name", "keys": ["alternativePersonName"]},
                    {"label": "Alternative Person Designation", "keys": ["alternativePersonDesignation"]},
                    {"label": "Alternative Person Email", "keys": ["alternativePersonEmail"]},
                    {"label": "Alternative Person Mobile", "keys": ["alternativePersonMobile"]},
                    {"label": "Date", "keys": ["alternativePersonDate"]},
                    {"label": "Person Place", "keys": ["alternativePersonPlace"]},
                ],
            },
        ],
    },
    {
        "title": "6. Sharp & Tannan Team",
        "cards": [
            {
                "title": "Sharp & Tannan Team",
                "fields": [
                    {"label": "Inco Terms", "keys": ["incoTerms", "IncoTerms", "Incoterms"]},
                    {"label": "Payment Terms", "keys": ["paymentTerms", "PaymentTerms"]},
                    {"label": "Vendor Type", "keys": ["vendorType", "VendorType"]},
                    {"label": "Sharp & Tannan Approver HOD", "keys": ["rubaminApproverHod", "RubaminApproverHod", "hodEmail", "HodEmail"]},
                ],
            },
        ],
    },
]

PDF_BROWSER_PATH_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def _is_meaningful_pdf_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str) and value.strip() == "":
        return False
    return True


def _pick_pdf_field_value(record: dict[str, Any], keys: list[str], fallback: Any = "") -> Any:
    for key in keys:
        if key not in record:
            continue
        value = record.get(key)
        if _is_meaningful_pdf_value(value):
            return value
    for key in keys:
        if key in record:
            return record.get(key)
    return fallback


def _format_pdf_field_value(value: Any) -> str:
    if value is None:
        return "-"
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, (list, tuple, set)):
        items = [str(item or "").strip() for item in value if str(item or "").strip()]
        return ", ".join(items) if items else "-"
    if isinstance(value, dict):
        try:
            return json.dumps(value, ensure_ascii=False)
        except Exception:
            return str(value)
    text = str(value).strip()
    return text or "-"


def _extract_document_display_value_from_record(record: dict[str, Any], keys: list[str]) -> str:
    raw_value = _pick_pdf_field_value(record, keys, "")
    value = str(raw_value or "").strip()
    if not value:
        return "No Document"

    filename = _extract_document_filename(value)
    if filename:
        return filename

    direct_name = Path(value.split("?", 1)[0]).name.strip()
    if direct_name:
        return direct_name

    return "Uploaded Document"


def _build_vendor_form_field_html(record: dict[str, Any], field: dict[str, Any]) -> str:
    label = str(field.get("label") or "").strip() or "Field"
    keys = [str(item or "").strip() for item in (field.get("keys") or []) if str(item or "").strip()]
    is_doc = bool(field.get("doc"))
    is_multiline = bool(field.get("multiline"))
    is_full = bool(field.get("full")) or is_doc or is_multiline

    if is_doc:
        value_text = _extract_document_display_value_from_record(record, keys)
    else:
        value_text = _format_pdf_field_value(_pick_pdf_field_value(record, keys, ""))

    row_class = "row full" if is_full else "row"
    value_class = "value multiline" if is_multiline else "value"
    return (
        f'<div class="{row_class}">'
        f"<label>{html_escape(label)}</label>"
        f'<div class="{value_class}">{html_escape(value_text)}</div>'
        "</div>"
    )


def _normalize_pdf_lookup_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _pick_pdf_field_value_by_fragments(record: dict[str, Any], fragments: list[str]) -> Any:
    normalized_fragments = [str(item or "").strip().lower() for item in fragments if str(item or "").strip()]
    if not normalized_fragments or not isinstance(record, dict):
        return ""

    for key, value in record.items():
        normalized_key = _normalize_pdf_lookup_key(key)
        if not normalized_key:
            continue
        if all(fragment in normalized_key for fragment in normalized_fragments):
            if _is_meaningful_pdf_value(value):
                return value

    for key, value in record.items():
        normalized_key = _normalize_pdf_lookup_key(key)
        if not normalized_key:
            continue
        if all(fragment in normalized_key for fragment in normalized_fragments):
            return value

    return ""


def _build_vendor_form_static_field_html(label: str, value: Any) -> str:
    text = _format_pdf_field_value(value)
    return (
        '<div class="row">'
        f"<label>{html_escape(label)}</label>"
        f'<div class="value">{html_escape(text)}</div>'
        "</div>"
    )


def _build_vendor_form_pdf_html(record_id: str, record: dict[str, Any]) -> str:
    vendor_name = _extract_vendor_name_for_history_export(record) or "-"
    status = _format_pdf_field_value(
        _pick_pdf_field_value(record, ["approverStatus", "ApproverStatus", "status", "Status"], "")
    )

    section_html_parts: list[str] = []
    for section in PDF_VENDOR_FORM_SECTIONS:
        section_title_text = str(section.get("title") or "").strip()
        if section_title_text.lower().startswith("6. rubamin and it") or section_title_text.lower().startswith("6. sharp"):
            continue
        card_html_parts: list[str] = []
        for card in section.get("cards", []):
            visible_fields_html_parts: list[str] = []
            for field in card.get("fields", []):
                # Per requirement: document/upload fields should not appear in Vendor Form PDF.
                if bool(field.get("doc")):
                    continue
                visible_fields_html_parts.append(_build_vendor_form_field_html(record, field))

            fields_html = "".join(visible_fields_html_parts).strip()
            if not fields_html:
                continue

            card_html_parts.append(
                "<div class=\"subcard\">"
                f"<h3>{html_escape(str(card.get('title') or ''))}</h3>"
                f"<div class=\"grid\">{fields_html}</div>"
                "</div>"
            )
        if not card_html_parts:
            continue
        section_html_parts.append(f"<h2>{html_escape(str(section.get('title') or ''))}</h2>")
        section_html_parts.extend(card_html_parts)

    # Always render Sharp & Tannan Team block explicitly so these 4 fields are visible in PDF.
    rubamin_team_values = {
        "Inco Terms": _pick_pdf_field_value(record, ["incoTerms", "IncoTerms", "Incoterms", "inco terms"], ""),
        "Payment Terms": _pick_pdf_field_value(record, ["paymentTerms", "PaymentTerms", "payment terms"], ""),
        "Vendor Type": _pick_pdf_field_value(record, ["vendorType", "VendorType", "vendor type"], ""),
        "Sharp & Tannan Approver HOD": _pick_pdf_field_value(
            record,
            ["rubaminApproverHod", "RubaminApproverHod", "RubaminApprovalHod", "hodEmail", "HodEmail"],
            "",
        ),
    }
    if not _is_meaningful_pdf_value(rubamin_team_values["Inco Terms"]):
        rubamin_team_values["Inco Terms"] = _pick_pdf_field_value_by_fragments(record, ["inco", "term"])
    if not _is_meaningful_pdf_value(rubamin_team_values["Payment Terms"]):
        rubamin_team_values["Payment Terms"] = _pick_pdf_field_value_by_fragments(record, ["payment", "term"])
    if not _is_meaningful_pdf_value(rubamin_team_values["Vendor Type"]):
        rubamin_team_values["Vendor Type"] = _pick_pdf_field_value_by_fragments(record, ["vendor", "type"])
    if not _is_meaningful_pdf_value(rubamin_team_values["Sharp & Tannan Approver HOD"]):
        rubamin_team_values["Sharp & Tannan Approver HOD"] = (
            _pick_pdf_field_value_by_fragments(record, ["rubamin", "approver", "hod"])
            or _pick_pdf_field_value_by_fragments(record, ["hod", "email"])
        )

    rubamin_team_fields_html = "".join(
        _build_vendor_form_static_field_html(label, value)
        for label, value in rubamin_team_values.items()
    )
    section_html_parts.append("<h2>6. Sharp &amp; Tannan Team</h2>")
    section_html_parts.append(
        '<div class="subcard">'
        "<h3>Sharp &amp; Tannan Team</h3>"
        f'<div class="grid team-grid">{rubamin_team_fields_html}</div>'
        "</div>"
    )

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Vendor Form - {html_escape(record_id)}</title>
  <style>
    @page {{ size: A4; margin: 10mm; }}
    * {{ box-sizing: border-box; }}
    body {{
      font-family: Arial, sans-serif;
      background: #ffffff;
      color: #1f2937;
      margin: 0;
      padding: 0;
      font-size: 11px;
    }}
    .wrap {{ width: 100%; }}
    .page-header {{
      border: 1px solid #dbeafe;
      border-radius: 10px;
      padding: 10px 12px;
      margin-bottom: 10px;
    }}
    .brand-text {{
      text-align: center;
      font-size: 13px;
      line-height: 1.35;
      color: #334155;
      font-weight: 700;
      margin-bottom: 8px;
    }}
    h1 {{
      margin: 0;
      font-size: 18px;
      color: #1e3a8a;
      text-align: center;
    }}
    .meta {{
      display: flex;
      gap: 6px;
      flex-wrap: wrap;
      margin-top: 8px;
      justify-content: center;
    }}
    .badge {{
      border: 1px solid #bfdbfe;
      background: #eff6ff;
      color: #1e3a8a;
      border-radius: 999px;
      padding: 4px 8px;
      font-weight: 700;
      font-size: 10px;
    }}
    h2 {{
      margin: 12px 0 6px;
      font-size: 13px;
      color: #1e3a8a;
      border-bottom: 1px solid #dbeafe;
      padding-bottom: 4px;
      page-break-after: avoid;
    }}
    .subcard {{
      border: 1px solid #dbeafe;
      background: #f8fbff;
      border-radius: 8px;
      padding: 8px;
      margin-bottom: 8px;
      page-break-inside: avoid;
    }}
    .subcard h3 {{
      margin: 0 0 8px;
      font-size: 12px;
      color: #1e3a8a;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 8px;
    }}
    .grid.team-grid {{
      grid-template-columns: repeat(4, minmax(0, 1fr));
    }}
    .row {{
      min-height: 44px;
      border: 1px solid #dbe3f1;
      border-radius: 6px;
      padding: 6px;
      background: #ffffff;
    }}
    .row.full {{ grid-column: 1 / -1; }}
    label {{
      display: block;
      font-size: 9px;
      font-weight: 700;
      color: #475569;
      margin-bottom: 4px;
      text-transform: uppercase;
      letter-spacing: .2px;
    }}
    .value {{
      font-size: 11px;
      color: #111827;
      line-height: 1.35;
      white-space: normal;
      word-break: break-word;
    }}
    .value.multiline {{
      white-space: pre-wrap;
      min-height: 36px;
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="page-header">
      <div class="brand-text">Sharp & Tannan Chartered Accountants</div>
      <h1>Vendor Registration Form</h1>
      <div class="meta">
        <span class="badge">Record ID: {html_escape(record_id)}</span>
        <span class="badge">Vendor: {html_escape(vendor_name)}</span>
        <span class="badge">Approver Status: {html_escape(status)}</span>
      </div>
    </div>
    {''.join(section_html_parts)}
  </div>
</body>
</html>"""


def _discover_pdf_browser_executables() -> list[str]:
    candidates: list[str] = []
    seen: set[str] = set()

    env_candidates = [
        os.getenv("PDF_BROWSER_PATH", "").strip(),
        os.getenv("MS_EDGE_PATH", "").strip(),
        os.getenv("CHROME_PATH", "").strip(),
        shutil.which("msedge") or "",
        shutil.which("chrome") or "",
        shutil.which("msedge.exe") or "",
        shutil.which("chrome.exe") or "",
    ]

    for item in env_candidates + PDF_BROWSER_PATH_CANDIDATES:
        path_text = str(item or "").strip()
        if not path_text:
            continue
        path_obj = Path(path_text)
        if not path_obj.exists() or not path_obj.is_file():
            continue
        resolved = str(path_obj.resolve())
        lowered = resolved.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        candidates.append(resolved)

    return candidates


def _render_html_to_pdf_with_browser(html_content: str) -> bytes:
    browser_paths = _discover_pdf_browser_executables()
    if not browser_paths:
        raise RuntimeError("Unable to generate layout PDF: no supported browser executable found (Edge/Chrome).")

    errors: list[str] = []
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        html_path = temp_path / "vendor_form_export.html"
        html_path.write_text(html_content, encoding="utf-8")

        for browser_path in browser_paths:
            for headless_flag in ("--headless=new", "--headless"):
                pdf_path = temp_path / "vendor_form_export.pdf"
                if pdf_path.exists():
                    pdf_path.unlink()
                user_data_dir = temp_path / f"profile_{hashlib.sha1((browser_path + headless_flag).encode('utf-8')).hexdigest()[:10]}"
                user_data_dir.mkdir(parents=True, exist_ok=True)

                cmd = [
                    browser_path,
                    headless_flag,
                    "--disable-gpu",
                    "--no-first-run",
                    "--no-default-browser-check",
                    f"--user-data-dir={user_data_dir}",
                    "--print-to-pdf-no-header",
                    f"--print-to-pdf={pdf_path}",
                    html_path.resolve().as_uri(),
                ]

                try:
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        timeout=120,
                        check=False,
                    )
                except Exception as exc:
                    errors.append(f"{Path(browser_path).name} {headless_flag}: {exc}")
                    continue

                if result.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0:
                    return pdf_path.read_bytes()

                stderr_text = (result.stderr or "").strip().replace("\n", " ")
                errors.append(
                    f"{Path(browser_path).name} {headless_flag}: rc={result.returncode} {stderr_text[:240]}"
                )

    error_preview = "; ".join(errors[:3]) if errors else "unknown error"
    raise RuntimeError(f"Unable to generate layout PDF. {error_preview}")


def _build_vendor_form_layout_pdf_bytes(record_id: str, record: dict[str, Any]) -> bytes:
    html_content = _build_vendor_form_pdf_html(record_id, record)
    return _render_html_to_pdf_with_browser(html_content)


def _build_single_record_pdf_bytes(record_id: str, record: dict[str, Any]) -> bytes:
    return _build_vendor_form_layout_pdf_bytes(record_id, record)


def _insert_paragraph_before_docx_table(table: Any, text: str) -> DocxParagraph:
    paragraph_element = OxmlElement("w:p")
    table._tbl.addprevious(paragraph_element)
    paragraph = DocxParagraph(paragraph_element, table._parent)
    run = paragraph.add_run(text)
    run.bold = True
    return paragraph


def _apply_information_furnished_by_to_code_of_conduct(
    document: DocxDocument,
    information_furnished_by: str,
) -> None:
    furnished_by = _format_pdf_field_value(information_furnished_by)
    if not furnished_by:
        return

    line_text = f"Information Furnished By: {furnished_by}"
    if document.tables:
        _insert_paragraph_before_docx_table(document.tables[-1], line_text)
        return

    paragraph = document.add_paragraph()
    run = paragraph.add_run(line_text)
    run.bold = True


def _build_code_of_conduct_html_from_docx(docx_path: Path) -> str:
    document = DocxDocument(str(docx_path))
    content_parts: list[str] = []

    for index, paragraph in enumerate(document.paragraphs):
        text = str(paragraph.text or "").strip()
        if not text:
            continue
        escaped_text = html_escape(text)
        if index == 0:
            content_parts.append(f"<h1>{escaped_text}</h1>")
        elif text.lower() in {
            "introduction",
            "guiding principles for suppliers and contractors",
            "labor and human rights",
            "business ethics and transparency",
            "occupational health & safety",
            "environmental consciousness",
        }:
            content_parts.append(f"<h2>{escaped_text}</h2>")
        else:
            content_parts.append(f"<p>{escaped_text}</p>")

    if document.tables:
        table = document.tables[-1]
        table_rows: list[str] = []
        for row in table.rows:
            cell_html = "".join(f"<td>{html_escape(str(cell.text or '').strip())}</td>" for cell in row.cells)
            if cell_html:
                table_rows.append(f"<tr>{cell_html}</tr>")
        if table_rows:
            content_parts.append(
                "<table class=\"sign-table\">"
                "<tbody>"
                + "".join(table_rows)
                + "</tbody></table>"
            )

    return (
        "<!doctype html><html><head><meta charset='utf-8'>"
        "<style>"
        "body{font-family:Arial,sans-serif;color:#111827;margin:32px;line-height:1.55;font-size:13px;}"
        "h1{font-size:22px;margin:0 0 18px;color:#0f172a;}"
        "h2{font-size:16px;margin:20px 0 10px;color:#1e3a8a;}"
        "p{margin:0 0 10px;}"
        ".sign-table{width:100%;border-collapse:collapse;margin-top:24px;}"
        ".sign-table td{border:1px solid #cbd5e1;padding:10px;vertical-align:top;font-weight:600;}"
        "</style></head><body>"
        + "".join(content_parts)
        + "</body></html>"
    )


def _read_text_file_with_fallback_encodings(file_path: Path) -> str:
    raw_bytes = file_path.read_bytes()
    for encoding in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("utf-8", errors="ignore")


def _convert_docx_to_preview_html_with_word(docx_path: Path) -> Path:
    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client as win32com  # type: ignore[import-untyped]
    except Exception as exc:
        raise RuntimeError(f"Word COM dependencies are unavailable: {exc}") from exc

    if not docx_path.exists() or not docx_path.is_file():
        raise RuntimeError(f"Code of Conduct DOCX not found: {docx_path}")

    pythoncom.CoInitialize()
    word_app = None
    document = None
    try:
        html_path = docx_path.with_suffix(".html")
        html_assets_dir = html_path.with_name(f"{html_path.stem}_files")
        if html_path.exists():
            html_path.unlink()
        if html_assets_dir.exists():
            shutil.rmtree(html_assets_dir, ignore_errors=True)

        word_app = win32com.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0
        document = word_app.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        try:
            document.SaveAs2(str(html_path), FileFormat=8)
        except Exception:
            document.SaveAs(str(html_path), FileFormat=8)

        if not html_path.exists() or html_path.stat().st_size <= 0:
            raise RuntimeError("Word did not generate an HTML output")
        return html_path
    finally:
        try:
            if document is not None:
                document.Close(False)
        except Exception:
            pass
        try:
            if word_app is not None:
                word_app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _resolve_code_of_conduct_preview_asset_path(reference: str, html_path: Path) -> Path | None:
    raw_reference = str(reference or "").strip().strip("\"'")
    if not raw_reference:
        return None

    lowered_reference = raw_reference.lower()
    if lowered_reference.startswith(("data:", "http:", "https:", "mailto:", "javascript:", "#")):
        return None

    normalized_reference = unquote(raw_reference).replace("\\", "/")
    assets_dir = html_path.with_name(f"{html_path.stem}_files")
    if not assets_dir.exists() or not assets_dir.is_dir():
        return None

    marker = f"{html_path.stem}_files/"
    if marker in normalized_reference:
        relative_part = normalized_reference.split(marker, 1)[1]
    else:
        cleaned_reference = normalized_reference.lstrip("./")
        if "/" in cleaned_reference:
            prefix, _, remainder = cleaned_reference.partition("/")
            if prefix.endswith("_files"):
                relative_part = remainder
            else:
                return None
        else:
            relative_part = cleaned_reference

    asset_path = (assets_dir / relative_part).resolve()
    try:
        asset_path.relative_to(assets_dir.resolve())
    except ValueError:
        return None

    if not asset_path.exists() or not asset_path.is_file():
        return None
    return asset_path


def _inline_code_of_conduct_preview_assets(html_content: str, html_path: Path) -> str:
    output = str(html_content or "")

    def _replace_link_tag(match: re.Match[str]) -> str:
        tag_text = match.group(0)
        href_match = re.search(r'href=["\']([^"\']+)["\']', tag_text, flags=re.IGNORECASE)
        if not href_match:
            return tag_text

        href_value = href_match.group(1)
        normalized_href_value = unquote(str(href_value or "")).replace("\\", "/")
        asset_path = _resolve_code_of_conduct_preview_asset_path(href_value, html_path)
        if asset_path is None:
            if "_files/" in normalized_href_value:
                return ""
            return tag_text

        if asset_path.suffix.lower() == ".css":
            return f"<style>\n{_read_text_file_with_fallback_encodings(asset_path)}\n</style>"
        return ""

    output = re.sub(r"(?is)<link\b[^>]*>", _replace_link_tag, output)

    def _replace_asset_attribute(match: re.Match[str]) -> str:
        attribute_name = str(match.group(1) or "")
        reference = str(match.group(2) or "")
        asset_path = _resolve_code_of_conduct_preview_asset_path(reference, html_path)
        if asset_path is None:
            return match.group(0)

        suffix = asset_path.suffix.lower()
        if suffix not in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}:
            if attribute_name.lower() == "href":
                return f'{attribute_name}="#"'
            return match.group(0)

        mime_type = str(mimetypes.guess_type(str(asset_path))[0] or "application/octet-stream")
        encoded = base64.b64encode(asset_path.read_bytes()).decode("ascii")
        return f'{attribute_name}="data:{mime_type};base64,{encoded}"'

    output = re.sub(
        r'(?i)\b(src|href)=["\']([^"\']+)["\']',
        _replace_asset_attribute,
        output,
    )

    def _replace_css_url(match: re.Match[str]) -> str:
        reference = str(match.group(2) or "")
        asset_path = _resolve_code_of_conduct_preview_asset_path(reference, html_path)
        if asset_path is None:
            normalized_reference = unquote(reference).replace("\\", "/")
            if "_files/" in normalized_reference:
                return "none"
            return match.group(0)

        suffix = asset_path.suffix.lower()
        if suffix not in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}:
            return "none"

        mime_type = str(mimetypes.guess_type(str(asset_path))[0] or "application/octet-stream")
        encoded = base64.b64encode(asset_path.read_bytes()).decode("ascii")
        return f'url("data:{mime_type};base64,{encoded}")'

    output = re.sub(
        r'(?i)url\((["\']?)([^"\')]+)\1\)',
        _replace_css_url,
        output,
    )

    output = re.sub(
        r"(?is)<xml>.*?</xml>",
        "",
        output,
    )
    return output


def _wrap_code_of_conduct_preview_html(html_content: str) -> str:
    raw_output = str(html_content or "")

    html_open_match = re.search(r"<html\b(?P<attrs>[^>]*)>", raw_output, flags=re.IGNORECASE)
    head_match = re.search(r"<head\b[^>]*>(?P<content>.*?)</head>", raw_output, flags=re.IGNORECASE | re.DOTALL)
    body_open_match = re.search(r"<body\b(?P<attrs>[^>]*)>", raw_output, flags=re.IGNORECASE)
    body_match = re.search(r"<body\b[^>]*>(?P<content>.*?)</body>", raw_output, flags=re.IGNORECASE | re.DOTALL)

    html_attrs = html_open_match.group("attrs") if html_open_match else ""
    head_content = head_match.group("content") if head_match else ""
    body_attrs = body_open_match.group("attrs") if body_open_match else ""
    body_content = body_match.group("content") if body_match else raw_output

    preview_styles = """
<style>
html { background: #e5e7eb; }
body {
  margin: 0 !important;
  min-height: 100vh;
  background: #e5e7eb !important;
}
.code-of-conduct-preview-root {
  box-sizing: border-box;
  min-height: 100vh;
  padding: 24px 18px 36px;
}
.code-of-conduct-preview-page {
  box-sizing: border-box;
  width: min(100%, 920px);
  margin: 0 auto;
  background: #fff;
  border: 1px solid #d6deeb;
  box-shadow: 0 12px 28px rgba(15, 23, 42, 0.10);
  padding: 34px 42px;
  overflow: auto;
}
.code-of-conduct-preview-page img {
  max-width: 100% !important;
  height: auto !important;
}
.code-of-conduct-preview-page table {
  max-width: 100% !important;
}
.code-of-conduct-preview-page p,
.code-of-conduct-preview-page div,
.code-of-conduct-preview-page li,
.code-of-conduct-preview-page td,
.code-of-conduct-preview-page th {
  overflow-wrap: break-word;
}
@media (max-width: 900px) {
  .code-of-conduct-preview-root {
    padding: 16px 10px 24px;
  }
  .code-of-conduct-preview-page {
    padding: 22px 18px;
  }
}
</style>
"""

    return (
        "<!doctype html>"
        f"<html{html_attrs}>"
        "<head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        + head_content
        + preview_styles
        + "</head>"
        f"<body{body_attrs}>"
        "<div class='code-of-conduct-preview-root'>"
        f"<div class='code-of-conduct-preview-page'>{body_content}</div>"
        "</div></body></html>"
    )


def _convert_docx_to_pdf_with_word(docx_path: Path) -> bytes:
    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client as win32com  # type: ignore[import-untyped]
    except Exception as exc:
        raise RuntimeError(f"Word COM dependencies are unavailable: {exc}") from exc

    if not docx_path.exists() or not docx_path.is_file():
        raise RuntimeError(f"Code of Conduct DOCX not found: {docx_path}")

    pythoncom.CoInitialize()
    word_app = None
    document = None
    try:
        pdf_path = docx_path.with_suffix(".pdf")
        if pdf_path.exists():
            pdf_path.unlink()

        word_app = win32com.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0
        document = word_app.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        document.ExportAsFixedFormat(str(pdf_path), 17)
        if not pdf_path.exists() or pdf_path.stat().st_size <= 0:
            raise RuntimeError("Word did not generate a PDF output")
        return pdf_path.read_bytes()
    finally:
        try:
            if document is not None:
                document.Close(False)
        except Exception:
            pass
        try:
            if word_app is not None:
                word_app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _build_code_of_conduct_pdf_bytes(information_furnished_by: str = "") -> bytes:
    if not CODE_OF_CONDUCT_TEMPLATE_PATH.exists():
        raise RuntimeError(f"Code of Conduct template not found: {CODE_OF_CONDUCT_TEMPLATE_PATH}")

    with tempfile.TemporaryDirectory(prefix="code_of_conduct_") as temp_dir:
        working_docx_path = Path(temp_dir) / CODE_OF_CONDUCT_TEMPLATE_PATH.name
        shutil.copy2(CODE_OF_CONDUCT_TEMPLATE_PATH, working_docx_path)

        document = DocxDocument(str(working_docx_path))
        _apply_information_furnished_by_to_code_of_conduct(document, information_furnished_by)
        document.save(str(working_docx_path))

        try:
            with CODE_OF_CONDUCT_RENDER_LOCK:
                return _convert_docx_to_pdf_with_word(working_docx_path)
        except Exception:
            html_content = _build_code_of_conduct_html_from_docx(working_docx_path)
            return _render_html_to_pdf_with_browser(html_content)


def _build_code_of_conduct_view_html() -> str:
    if not CODE_OF_CONDUCT_TEMPLATE_PATH.exists():
        raise RuntimeError(f"Code of Conduct template not found: {CODE_OF_CONDUCT_TEMPLATE_PATH}")

    with tempfile.TemporaryDirectory(prefix="code_of_conduct_view_") as temp_dir:
        working_docx_path = Path(temp_dir) / CODE_OF_CONDUCT_TEMPLATE_PATH.name
        shutil.copy2(CODE_OF_CONDUCT_TEMPLATE_PATH, working_docx_path)
        try:
            with CODE_OF_CONDUCT_RENDER_LOCK:
                html_path = _convert_docx_to_preview_html_with_word(working_docx_path)
            raw_html = _read_text_file_with_fallback_encodings(html_path)
            html_content = _inline_code_of_conduct_preview_assets(raw_html, html_path)
        except Exception:
            html_content = _build_code_of_conduct_html_from_docx(working_docx_path)
        return _wrap_code_of_conduct_preview_html(html_content)


def _extract_code_of_conduct_information_furnished_by(
    local_record: dict[str, Any] | None,
    db_record: dict[str, Any] | None,
) -> str:
    return str(
        _first_existing_record_value(
            local_record,
            db_record,
            "contactPersonName",
            "ContactPerson",
            "FurnishBy",
            "furnishBy",
            "vendorName",
            "Vendorname",
            "VendorName",
        )
        or ""
    ).strip()


def _store_vendor_code_of_conduct_pdf(record_id: str) -> None:
    normalized_record_id = _normalize_record_id(record_id)
    if not normalized_record_id:
        return

    local_record = find_local_vendor_record_by_record_id(normalized_record_id)
    db_record = get_vendor_record_by_record_id(normalized_record_id)
    if not local_record and not db_record:
        return

    information_furnished_by = _extract_code_of_conduct_information_furnished_by(local_record, db_record)
    pdf_bytes = _build_code_of_conduct_pdf_bytes(information_furnished_by)
    sanitized_record_id = _sanitize_uploaded_file_segment(normalized_record_id, "record")
    file_name = f"Supplier_and_Contractor_Code_of_Conduct_{sanitized_record_id}.pdf"
    document_url = _store_uploaded_document(
        normalized_record_id,
        CODE_OF_CONDUCT_DOCUMENT_FIELD,
        file_name,
        pdf_bytes,
        "application/pdf",
    )

    sql_updates: dict[str, Any] = {
        CODE_OF_CONDUCT_DOCUMENT_FIELD: document_url,
        f"{CODE_OF_CONDUCT_DOCUMENT_FIELD}_FileName": file_name,
        f"{CODE_OF_CONDUCT_DOCUMENT_FIELD}_ContentType": "application/pdf",
        f"{CODE_OF_CONDUCT_DOCUMENT_FIELD}_Data": pdf_bytes,
    }
    ensure_vendor_columns(list(sql_updates.keys()))
    update_local_vendor_record_by_record_id(normalized_record_id, sql_updates)
    if not update_vendor_record_by_record_id(normalized_record_id, sql_updates):
        raise RuntimeError("Failed to save Code of Conduct PDF against vendor record")


def _resolve_vendor_document_for_export(
    record_id: str,
    field_name: str,
    local_record: dict[str, Any] | None,
    db_record: dict[str, Any] | None,
) -> dict[str, Any] | None:
    candidate_keys = _get_document_candidate_keys(field_name)
    if not candidate_keys:
        return None

    canonical_field_name = candidate_keys[0]
    stored_document = get_vendor_document_content_by_record_id(record_id, candidate_keys)
    if stored_document:
        file_name = _extract_document_filename(stored_document.get("filename")) or canonical_field_name
        content = stored_document.get("content") or b""
        if content:
            content_type = str(stored_document.get("content_type") or "").strip() or (
                mimetypes.guess_type(file_name)[0] or "application/octet-stream"
            )
            return {
                "filename": file_name,
                "content": content,
                "content_type": content_type,
            }

    raw_reference_candidates: list[str] = []
    seen_references: set[str] = set()
    for record in (db_record, local_record):
        if not isinstance(record, dict):
            continue
        for key in candidate_keys:
            value = str(record.get(key) or "").strip()
            if not value:
                continue
            dedupe_key = value.lower()
            if dedupe_key in seen_references:
                continue
            seen_references.add(dedupe_key)
            raw_reference_candidates.append(value)

    extracted_file_names: list[str] = []
    seen_file_names: set[str] = set()
    for raw_reference in raw_reference_candidates:
        file_name = _extract_document_filename(raw_reference)
        if not file_name:
            continue
        dedupe_key = file_name.lower()
        if dedupe_key in seen_file_names:
            continue
        seen_file_names.add(dedupe_key)
        extracted_file_names.append(file_name)

    field_name_candidates = [canonical_field_name, *candidate_keys[1:]]
    for field_key in field_name_candidates:
        for extracted_name in extracted_file_names:
            disk_path = _find_uploaded_file_for_record_field(record_id, field_key, extracted_name)
            if disk_path is not None and disk_path.exists() and disk_path.is_file():
                return {
                    "filename": disk_path.name,
                    "content": disk_path.read_bytes(),
                    "content_type": mimetypes.guess_type(str(disk_path))[0] or "application/octet-stream",
                }
        disk_path = _find_uploaded_file_for_record_field(record_id, field_key)
        if disk_path is not None and disk_path.exists() and disk_path.is_file():
            return {
                "filename": disk_path.name,
                "content": disk_path.read_bytes(),
                "content_type": mimetypes.guess_type(str(disk_path))[0] or "application/octet-stream",
            }

    for raw_reference in raw_reference_candidates:
        reference_value = str(raw_reference or "").strip()
        if not reference_value:
            continue
        if reference_value.lower().startswith("http://") or reference_value.lower().startswith("https://"):
            continue

        if reference_value.startswith("/mock-files/"):
            reference_value = reference_value.split("/mock-files/", 1)[1]

        storage_key, normalized_name = _normalize_mock_file_reference(reference_value)
        if not normalized_name:
            continue

        cached = mock_files.get(storage_key) or mock_files.get(normalized_name)
        if cached:
            content = cached.get("content") or b""
            if content:
                content_type = str(cached.get("content_type") or "").strip() or (
                    mimetypes.guess_type(normalized_name)[0] or "application/octet-stream"
                )
                return {
                    "filename": normalized_name,
                    "content": content,
                    "content_type": content_type,
                }

        disk_path = _resolve_uploaded_file_disk_path(storage_key)
        if disk_path is not None and disk_path.exists() and disk_path.is_file():
            return {
                "filename": disk_path.name,
                "content": disk_path.read_bytes(),
                "content_type": mimetypes.guess_type(str(disk_path))[0] or "application/octet-stream",
            }

    for extracted_name in extracted_file_names:
        disk_path = _find_uploaded_file_by_name(extracted_name)
        if disk_path is not None and disk_path.exists() and disk_path.is_file():
            return {
                "filename": disk_path.name,
                "content": disk_path.read_bytes(),
                "content_type": mimetypes.guess_type(str(disk_path))[0] or "application/octet-stream",
            }

        stored = get_vendor_document_content_by_filename(extracted_name)
        if stored:
            content = stored.get("content") or b""
            if content:
                content_type = str(stored.get("content_type") or "").strip() or (
                    mimetypes.guess_type(extracted_name)[0] or "application/octet-stream"
                )
                return {
                    "filename": extracted_name,
                    "content": content,
                    "content_type": content_type,
                }

    return None


def _build_unique_zip_entry_name(file_name: str, used_names: set[str], fallback_prefix: str) -> str:
    base_name = Path(str(file_name or "")).name.strip()
    if not base_name:
        base_name = f"{fallback_prefix}.bin"
    base_name = re.sub(r"[\\/:*?\"<>|]+", "-", base_name).strip(" .")
    if not base_name:
        base_name = f"{fallback_prefix}.bin"

    stem = Path(base_name).stem or fallback_prefix
    suffix = Path(base_name).suffix
    candidate = f"{stem}{suffix}"
    counter = 2
    while candidate.lower() in used_names:
        candidate = f"{stem}_{counter}{suffix}"
        counter += 1
    used_names.add(candidate.lower())
    return candidate


@app.get("/api/vendor/history-records/export")
def export_vendor_history_record(
    recordId: str,
    exportType: str,
    user: dict[str, Any] = Depends(get_current_user),
):
    normalized_export_type = str(exportType or "").strip().lower()
    if normalized_export_type not in {"excel", "pdf", "documents"}:
        raise HTTPException(status_code=400, detail="exportType must be excel, pdf, or documents")

    record_id, merged_record, local_record, db_record = _get_history_record_export_context(user, recordId)
    filename_base = _build_record_export_filename_base(record_id, merged_record)

    if normalized_export_type == "excel":
        excel_bytes = _build_single_record_excel_bytes(merged_record)
        headers_map = {"Content-Disposition": f'attachment; filename="{filename_base}.xlsx"'}
        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers_map,
        )

    if normalized_export_type == "pdf":
        try:
            pdf_bytes = _build_single_record_pdf_bytes(record_id, merged_record)
        except RuntimeError as exc:
            raise HTTPException(status_code=500, detail=str(exc))
        headers_map = {"Content-Disposition": f'attachment; filename="{filename_base}.pdf"'}
        return Response(content=pdf_bytes, media_type="application/pdf", headers=headers_map)

    document_fields: list[str] = list(DOCUMENT_FIELDS)
    for field_name in IMPORT_FORM_UPLOAD_TO_COLUMN.keys():
        if field_name not in document_fields:
            document_fields.append(field_name)

    collected_documents: list[dict[str, Any]] = []
    seen_content_hashes: set[str] = set()
    for index, field_name in enumerate(document_fields, start=1):
        payload = _resolve_vendor_document_for_export(record_id, field_name, local_record, db_record)
        if not payload:
            continue
        content = payload.get("content") or b""
        if not content:
            continue
        content_hash = hashlib.sha1(content).hexdigest()
        if content_hash in seen_content_hashes:
            continue
        seen_content_hashes.add(content_hash)
        collected_documents.append(
            {
                "fallback_name": f"document_{index}",
                "filename": str(payload.get("filename") or ""),
                "content": content,
            }
        )

    if not collected_documents:
        raise HTTPException(status_code=404, detail="No uploaded documents found for this record")

    output = BytesIO()
    used_entry_names: set[str] = set()
    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for item in collected_documents:
            entry_name = _build_unique_zip_entry_name(
                file_name=item.get("filename") or "",
                used_names=used_entry_names,
                fallback_prefix=str(item.get("fallback_name") or "document"),
            )
            zip_file.writestr(entry_name, item.get("content") or b"")
    output.seek(0)

    headers_map = {"Content-Disposition": f'attachment; filename="{filename_base}.zip"'}
    return StreamingResponse(output, media_type="application/zip", headers=headers_map)


@app.get("/api/vendor/history-records/download")
def download_vendor_history_records_excel(
    user: dict[str, Any] = Depends(get_current_user),
    recordId: str | None = None,
) -> Response:
    normalized_role = normalize_role(user.get("role"))
    user_email = str(user.get("email") or "").strip().lower()
    if normalized_role != "admin" and not user_email:
        raise HTTPException(status_code=403, detail="Unable to identify logged-in user")

    requested_record_id = _normalize_record_id(recordId)
    if recordId is not None and not requested_record_id:
        raise HTTPException(status_code=400, detail="recordId is invalid")

    summary_rows = get_vendor_history_records(user)
    summary_rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)

    record_ids: list[str] = []
    seen_record_ids: set[str] = set()
    record_id_by_key: dict[str, str] = {}
    summary_row_by_record_id: dict[str, dict[str, Any]] = {}
    for db_row in summary_rows:
        record_id = _normalize_record_id(db_row.get("record"))
        if not record_id:
            continue
        record_key = record_id.lower()
        if record_key in seen_record_ids:
            continue
        seen_record_ids.add(record_key)
        record_ids.append(record_id)
        record_id_by_key[record_key] = record_id
        summary_row_by_record_id[record_key] = db_row

    if requested_record_id:
        requested_key = requested_record_id.lower()
        if requested_key in record_id_by_key:
            record_ids = [record_id_by_key[requested_key]]
        else:
            requested_local_record = find_local_vendor_record_by_record_id(requested_record_id)
            requested_db_record = get_vendor_record_by_record_id(requested_record_id)
            if not requested_local_record and not requested_db_record:
                raise HTTPException(status_code=404, detail="Vendor record not found")
            _authorize_vendor_record_access(requested_local_record, requested_db_record, user)
            record_ids = [requested_record_id]

    export_rows: list[dict[str, Any]] = []
    headers: list[str] = []
    seen_headers: set[str] = set()

    for record_id in record_ids:
        record_key = record_id.lower()
        full_record = (
            get_vendor_record_by_record_id(record_id)
            or find_local_vendor_record_by_record_id(record_id)
            or {}
        )
        if not full_record:
            fallback_row = summary_row_by_record_id.get(record_key) or {}
            full_record = {
                "record": str(fallback_row.get("record") or record_id),
                "company": str(fallback_row.get("company") or ""),
                "category": str(fallback_row.get("category") or ""),
                "approverStatus": str(fallback_row.get("approverStatus") or ""),
                "name": str(fallback_row.get("name") or ""),
                "type": str(fallback_row.get("type") or ""),
                "createdAt": str(fallback_row.get("createdAt") or ""),
            }

        filtered_record: dict[str, Any] = {}
        for key, value in full_record.items():
            column_name = str(key or "").strip()
            if not column_name or _is_document_column_for_history_export(column_name):
                continue
            filtered_record[column_name] = value
            if column_name not in seen_headers:
                seen_headers.add(column_name)
                headers.append(column_name)

        if filtered_record:
            export_rows.append(filtered_record)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Vendor History"

    if not headers:
        worksheet.append(["No records found"])
        worksheet.column_dimensions["A"].width = 28
    else:
        worksheet.append(headers)
        column_widths = [len(str(header or "")) for header in headers]

        for row in export_rows:
            normalized_row = {
                _normalize_history_export_column_key(key): value
                for key, value in row.items()
                if str(key or "").strip()
            }
            row_values: list[Any] = []
            for idx, header in enumerate(headers):
                if header in row:
                    raw_value = row.get(header)
                else:
                    normalized_header = _normalize_history_export_column_key(header)
                    raw_value = normalized_row.get(normalized_header)
                cell_value = _coerce_excel_cell_value(raw_value)
                row_values.append(cell_value)

                text_length = len(str(cell_value or ""))
                if text_length > column_widths[idx]:
                    column_widths[idx] = text_length

            worksheet.append(row_values)

        for idx, width in enumerate(column_widths, start=1):
            worksheet.column_dimensions[get_column_letter(idx)].width = min(80, max(12, width + 2))
        worksheet.freeze_panes = "A2"

    output = BytesIO()
    workbook.save(output)
    workbook.close()
    excel_bytes = output.getvalue()

    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d_%H-%M-%S")
    if requested_record_id:
        filename = f"Vendor Record ({timestamp}).xlsx"
    else:
        filename = f"Vendor List ({timestamp}).xlsx"
    headers_map = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers_map,
    )


@app.post("/api/vendor/update-hod-email")
def update_vendor_hod_email(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    if not user_has_any_role(user, "user", "buyer"):
        raise HTTPException(status_code=403, detail="Only User role can update HOD email")

    record_id = _normalize_record_id(empty_to_none(body.get("recordId")))
    hod_email = str(empty_to_none(body.get("hodEmail")) or "").strip()
    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")
    if not hod_email:
        raise HTTPException(status_code=400, detail="hodEmail is required")
    if "@" not in hod_email or "." not in hod_email.split("@", 1)[-1]:
        raise HTTPException(status_code=400, detail="Invalid hodEmail format")

    local_record = find_local_vendor_record_by_record_id(record_id)
    db_record = get_vendor_record_by_record_id(record_id)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    user_email = str(user.get("email") or "").strip().lower()
    if not user_email:
        raise HTTPException(status_code=403, detail="Access denied for this record")

    assigned_to_user = (
        _record_is_assigned_to_email(local_record, user_email)
        or _record_is_assigned_to_email(db_record, user_email)
    )
    if not assigned_to_user:
        candidate_record_ids = {record_id.lower()}
        for row in get_vendor_records_by_assignee_email(user_email):
            row_record = _normalize_record_id(row.get("record")).lower()
            if row_record and row_record in candidate_record_ids:
                assigned_to_user = True
                break
    if not assigned_to_user:
        raise HTTPException(status_code=403, detail="Access denied for this record")

    current_status = str(
        _first_existing_record_value(
            db_record,
            local_record,
            "approverStatus",
            "ApproverStatus",
            "ApprovStatus",
            "inviteStatus",
            "status",
            "Status",
        )
        or ""
    ).strip().lower()
    if _is_pending_sap_code_creation_status(current_status):
        raise HTTPException(status_code=409, detail="HOD email cannot be updated after SAP stage")

    updates = {
        "rubaminApproverHod": hod_email,
        "RubaminApproverHod": hod_email,
        "Rubamin Approver HOD": hod_email,
        "RubaminApprovalHod": hod_email,
        "Rubamin Approval HOD": hod_email,
        "hodEmail": hod_email,
        "HodEmail": hod_email,
        "HODEmail": hod_email,
    }
    db_updated = update_vendor_record_by_record_id(record_id, updates)
    if db_record and not db_updated:
        raise HTTPException(status_code=500, detail="Vendor record found in database but update failed")

    local_updated = update_local_vendor_record_by_record_id(record_id, updates)
    if not local_updated and not db_updated:
        raise HTTPException(status_code=400, detail="Unable to update HOD email")

    return {
        "message": "HOD email updated successfully",
        "recordId": record_id,
        "hodEmail": hod_email,
    }


@app.get("/api/vendor/record")
def get_vendor_record_by_id(
    recordId: str,
    user: dict[str, Any] | None = Depends(get_optional_current_user),
) -> dict[str, Any]:
    record_id = _normalize_record_id(recordId)
    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")

    local_record = find_local_vendor_record_by_record_id(record_id)
    db_record = get_vendor_record_by_record_id(record_id)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    _authorize_vendor_record_access(local_record, db_record, user)

    merged_record: dict[str, Any] = {}
    if db_record:
        merged_record.update(db_record)
    if local_record:
        merged_record.update(local_record)
    db_status = str(
        _first_existing_record_value(
            db_record,
            None,
            "approverStatus",
            "ApproverStatus",
            "ApprovStatus",
            "inviteStatus",
            "status",
            "Status",
        )
        or ""
    ).strip()
    if db_status:
        merged_record["approverStatus"] = db_status
        merged_record["ApproverStatus"] = db_status
        if "ApprovStatus" in merged_record or (db_record and "ApprovStatus" in db_record):
            merged_record["ApprovStatus"] = db_status

    merged_record_id = _normalize_record_id(
        merged_record.get("recordId")
        or merged_record.get("RecordID")
        or merged_record.get("record")
        or merged_record.get("Record")
        or record_id
    )
    merged_record["recordId"] = merged_record_id

    pre_screen_answers: dict[str, str] = {}
    pre_screen_answers.update(_extract_prescreen_answers(db_record))
    pre_screen_answers.update(_extract_prescreen_answers(local_record))

    return {
        "recordId": merged_record_id,
        "vendor": merged_record,
        "preScreenAnswers": pre_screen_answers,
    }


@app.get("/api/vendor/gst-filing-table")
def get_vendor_gst_filing_table(
    recordId: str,
    financialYear: str | None = None,
    user: dict[str, Any] | None = Depends(get_optional_current_user),
) -> dict[str, Any]:
    record_id = _normalize_record_id(recordId)
    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")

    local_record = find_local_vendor_record_by_record_id(record_id)
    db_record = get_vendor_record_by_record_id(record_id)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    _authorize_vendor_record_access(local_record, db_record, user)

    approver_status = _first_existing_record_value(
        local_record,
        db_record,
        "approverStatus",
        "ApproverStatus",
        "status",
        "Status",
    )
    normalized_approver_status = _normalize_inline_status_text(approver_status)
    if normalized_approver_status not in {
        "pending for validator approval",
        "pending for validation approval",
        "pendign for validator approval",
    }:
        raise HTTPException(status_code=400, detail="GST filing table is available only for Pending For Validator Approval records")

    gst_status = _first_existing_record_value(
        local_record,
        db_record,
        "gstRegistrationStatus",
        "GSTStatus",
        "gstStatus",
    )
    if not _is_registered_gst_status(gst_status):
        raise HTTPException(status_code=400, detail="GST filing table is available only for GST registered records")

    gst_number = str(
        _first_existing_record_value(
            local_record,
            db_record,
            "gstNumber",
            "GSTIN",
            "gstin",
        )
        or ""
    ).strip().upper()
    if not gst_number:
        raise HTTPException(status_code=400, detail="GST number not found for this record")

    normalized_financial_year = _normalize_financial_year_value(financialYear)
    result = _run_gst_filing_table_extraction_inline(gst_number, financial_year=normalized_financial_year or None)
    if not isinstance(result, dict):
        raise HTTPException(status_code=500, detail="Unable to extract GST filing tables")

    return {
        "recordId": record_id,
        "gstNumber": gst_number,
        "financialYear": normalized_financial_year,
        "result": result,
    }


@app.post("/api/vendor/cancel")
def cancel_vendor_record(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    record_id = _normalize_record_id(empty_to_none(body.get("recordId")))
    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")

    local_record = find_local_vendor_record_by_record_id(record_id)
    db_record = get_vendor_record_by_record_id(record_id)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    normalized_role = normalize_role(user.get("role"))
    if normalized_role == "validator":
        raise HTTPException(status_code=403, detail="Validator role cannot cancel vendor")
    if normalized_role != "admin":
        user_email = str(user.get("email") or "").strip().lower()
        allowed_emails = {
            email
            for email in (
                _extract_record_email(local_record),
                _extract_record_email(db_record),
            )
            if email
        }
        if not user_email or not allowed_emails or user_email not in allowed_emails:
            raise HTTPException(status_code=403, detail="Access denied for this record")

    target_status = "Vendor Cancelled"
    local_updated = update_local_vendor_status_by_record_id(record_id, target_status)
    db_updated = update_vendor_approver_status_by_record_id(record_id, target_status)
    if not local_updated and not db_updated:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    return {
        "message": "Vendor cancelled successfully",
        "recordId": record_id,
        "approverStatus": target_status,
    }


@app.put("/api/vendor/update-record")
def update_vendor_record(
    body: dict[str, Any],
    request: Request,
    background_tasks: BackgroundTasks,
    user: dict[str, Any] | None = Depends(get_optional_current_user),
) -> dict[str, Any]:
    record_id = _normalize_record_id(empty_to_none(body.get("recordId")))
    updates_raw = body.get("updates")

    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")
    if not isinstance(updates_raw, dict) or not updates_raw:
        raise HTTPException(status_code=400, detail="updates are required")

    local_record = find_local_vendor_record_by_record_id(record_id)
    db_record = get_vendor_record_by_record_id(record_id)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    current_status = str(
        _first_existing_record_value(
            db_record,
            local_record,
            "approverStatus",
            "ApproverStatus",
            "ApprovStatus",
            "inviteStatus",
            "status",
            "Status",
        )
        or ""
    ).strip()
    resubmission_stage, resubmission_target_status = _resolve_resubmission_target_from_status(current_status)

    if user is not None:
        normalized_role = normalize_role(user.get("role"))
        if normalized_role == "validator":
            raise HTTPException(status_code=403, detail="Validator role cannot use update-record endpoint")
        if normalized_role != "admin":
            user_email = str(user.get("email") or "").strip().lower()
            allowed_emails = _collect_record_access_emails(local_record, db_record)
            if not user_email or not allowed_emails or user_email not in allowed_emails:
                raise HTTPException(status_code=403, detail="Access denied for this record")

    merged_record_for_lock: dict[str, Any] = {}
    if db_record:
        merged_record_for_lock.update(db_record)
    if local_record:
        merged_record_for_lock.update(local_record)
    locked_update_keys = get_validator_locked_update_keys(merged_record_for_lock)

    protected_keys = {"recordid", "record", "id"}
    updates: dict[str, Any] = {}
    for key, value in updates_raw.items():
        key_name = str(key or "").strip()
        if not key_name:
            continue
        key_name_lower = key_name.lower()
        if key_name_lower in protected_keys or key_name_lower in locked_update_keys:
            continue
        if isinstance(value, str):
            updates[key_name] = empty_to_none(value)
        else:
            updates[key_name] = value

    if resubmission_target_status:
        updates["approverStatus"] = resubmission_target_status
        updates["ApproverStatus"] = resubmission_target_status
        updates["rejectTo"] = None
        updates["Reject To"] = None

    sap_status_update_requested = _updates_request_sap_code_creation(updates)

    if not updates:
        raise HTTPException(status_code=400, detail="No updatable fields found")

    db_updated = update_vendor_record_by_record_id(record_id, updates)
    if db_record and not db_updated:
        raise HTTPException(status_code=500, detail="Vendor record found in database but update failed")

    local_updated = update_local_vendor_record_by_record_id(record_id, updates)
    if not local_updated and not db_updated:
        raise HTTPException(status_code=400, detail="Unable to update vendor record")

    refreshed_local = find_local_vendor_record_by_record_id(record_id)
    refreshed_db = get_vendor_record_by_record_id(record_id)
    merged_record: dict[str, Any] = {}
    if refreshed_db:
        merged_record.update(refreshed_db)
    if refreshed_local:
        merged_record.update(refreshed_local)
    refreshed_db_status = str(
        _first_existing_record_value(
            refreshed_db,
            None,
            "approverStatus",
            "ApproverStatus",
            "ApprovStatus",
            "inviteStatus",
            "status",
            "Status",
        )
        or ""
    ).strip()
    if refreshed_db_status:
        merged_record["approverStatus"] = refreshed_db_status
        merged_record["ApproverStatus"] = refreshed_db_status
        if "ApprovStatus" in merged_record or (refreshed_db and "ApprovStatus" in refreshed_db):
            merged_record["ApprovStatus"] = refreshed_db_status
    if not merged_record:
        merged_record = dict(updates)
        merged_record["recordId"] = record_id

    final_approver_status = str(
        _first_existing_record_value(
            merged_record,
            None,
            "approverStatus",
            "ApproverStatus",
            "ApprovStatus",
            "inviteStatus",
            "status",
            "Status",
        )
        or ""
    )

    approval_email_sent_count = 0
    approval_email_error: str | None = None
    if resubmission_target_status and resubmission_stage:
        dashboard_url = _build_public_registration_url(request, "/web/vendor-dashboard")
        vendor_name = (
            _extract_record_vendor_name(merged_record)
            or _extract_record_vendor_name(local_record)
            or _extract_record_vendor_name(db_record)
            or ""
        )
        submitted_on = str(
            _first_existing_record_value(
                merged_record,
                None,
                "updatedAt",
                "UpdatedAt",
                "createdAt",
                "CreatedAt",
            )
            or now_iso()
        )
        approval_email_sent_count, approval_email_error = _send_stage_approval_notifications(
            stage=resubmission_stage,
            record_id=record_id,
            vendor_name=vendor_name,
            submitted_on=submitted_on,
            approver_status=resubmission_target_status,
            dashboard_url=dashboard_url,
            merged_record=merged_record,
            local_record=local_record,
            db_record=db_record,
        )

    sap_sync_started, sap_sync_queued_count = _queue_sap_sync_for_pending_status(
        background_tasks,
        record_id,
        final_approver_status,
        force=sap_status_update_requested,
    )

    return {
        "message": (
            "Vendor record resubmitted successfully"
            if resubmission_target_status
            else "Vendor record updated successfully"
        ),
        "recordId": record_id,
        "approverStatus": (
            str(resubmission_target_status)
            if resubmission_target_status
            else final_approver_status
        ),
        "approvalEmailSent": approval_email_sent_count > 0,
        "approvalEmailSentCount": approval_email_sent_count,
        "approvalEmailError": approval_email_error,
        "resubmissionStage": resubmission_stage,
        "sapSyncStarted": sap_sync_started,
        "sapSyncQueuedCount": sap_sync_queued_count,
        "vendor": merged_record,
    }


@app.put("/api/vendor/update-record-with-files")
async def update_vendor_record_with_files(
    request: Request,
    background_tasks: BackgroundTasks,
    user: dict[str, Any] | None = Depends(get_optional_current_user),
) -> dict[str, Any]:
    def uses_filename_value(key_name: str) -> bool:
        key_lower = str(key_name or "").strip().lower()
        return (
            key_lower.endswith("filename")
            or key_lower.endswith("_filename")
            or key_lower.endswith("_name")
        )

    form_data = await request.form()

    record_id = _normalize_record_id(empty_to_none(form_data.get("recordId")))
    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")

    updates_raw_value = form_data.get("updates")
    updates_payload: dict[str, Any] = {}
    if isinstance(updates_raw_value, str) and updates_raw_value.strip():
        try:
            parsed_updates = json.loads(updates_raw_value)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="updates must be valid JSON")
        if not isinstance(parsed_updates, dict):
            raise HTTPException(status_code=400, detail="updates must be a JSON object")
        updates_payload = dict(parsed_updates)

    doc_source_map: dict[str, str] = {}
    doc_source_map_raw = form_data.get("docSourceMap")
    if isinstance(doc_source_map_raw, str) and doc_source_map_raw.strip():
        try:
            parsed_doc_source = json.loads(doc_source_map_raw)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="docSourceMap must be valid JSON")
        if not isinstance(parsed_doc_source, dict):
            raise HTTPException(status_code=400, detail="docSourceMap must be a JSON object")
        for key, value in parsed_doc_source.items():
            field_name = str(key or "").strip()
            source_key = str(value or "").strip()
            if field_name and source_key:
                doc_source_map[field_name] = source_key

    uploaded_doc_updates: dict[str, dict[str, Any]] = {}
    db_file_updates: dict[str, Any] = {}
    for key, value in form_data.multi_items():
        if not isinstance(value, (UploadFile, StarletteUploadFile)):
            continue
        key_name = str(key or "").strip()
        if not key_name.startswith("file_"):
            continue

        field_name = key_name[len("file_"):].strip()
        if not field_name or field_name not in DOCUMENT_FIELDS:
            continue

        raw_filename = str(value.filename or "").strip()
        normalized_filename = Path(raw_filename).name
        if not normalized_filename:
            continue

        file_bytes = await value.read()
        content = file_bytes or b""
        content_type = value.content_type or "application/octet-stream"
        doc_url = _store_uploaded_document(
            record_id,
            field_name,
            normalized_filename,
            content,
            content_type,
        )
        uploaded_doc_updates[field_name] = {
            "url": doc_url,
            "filename": normalized_filename,
        }
        db_file_updates[f"{field_name}_FileName"] = normalized_filename
        db_file_updates[f"{field_name}_ContentType"] = content_type
        db_file_updates[f"{field_name}_Data"] = content

    for field_name, file_meta in uploaded_doc_updates.items():
        doc_url = str(file_meta.get("url") or "")
        filename = str(file_meta.get("filename") or "")

        target_keys: set[str] = {field_name}
        mapped_source_key = str(doc_source_map.get(field_name) or "").strip()
        if mapped_source_key:
            target_keys.add(mapped_source_key)

        for key_name in target_keys:
            updates_payload[key_name] = filename if uses_filename_value(key_name) else doc_url

    if not updates_payload:
        raise HTTPException(status_code=400, detail="updates are required")

    result = update_vendor_record(
        {"recordId": record_id, "updates": updates_payload},
        request=request,
        background_tasks=background_tasks,
        user=user,
    )
    if db_file_updates:
        update_vendor_record_by_record_id(record_id, db_file_updates)
    return result


@app.post("/api/vendor/review-decision")
def submit_vendor_review_decision(
    body: dict[str, Any],
    request: Request,
    background_tasks: BackgroundTasks,
    user: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    record_id = _normalize_record_id(empty_to_none(body.get("recordId")))
    decision = str(empty_to_none(body.get("decision")) or "").strip().lower()
    remarks = str(empty_to_none(body.get("remarks")) or "").strip()
    team_raw = body.get("team")
    validations_raw = body.get("validations")
    tax_details_raw = body.get("taxDetails")
    reject_to = str(empty_to_none(body.get("rejectTo")) or "").strip()
    approval_mode = str(empty_to_none(body.get("approvalMode")) or "").strip().lower()

    if not record_id:
        raise HTTPException(status_code=400, detail="recordId is required")
    if decision not in {"approve", "reject"}:
        raise HTTPException(status_code=400, detail="decision must be Approve or Reject")
    if not remarks:
        raise HTTPException(status_code=400, detail="remarks are required")

    local_record = find_local_vendor_record_by_record_id(record_id)
    db_record = get_vendor_record_by_record_id(record_id)
    if not local_record and not db_record:
        raise HTTPException(status_code=404, detail="Vendor record not found")

    pending_user_statuses = {
        "pending for user approval",
        "pendign for user approval",
        "pending for buyer approval",
        "pending for data validation",
    }
    pending_validator_statuses = {
        "pending for validator approval",
        "pending for validation approval",
        "pendign for validator approval",
    }
    pending_hod_statuses = {
        "pending for hod approval",
        "pendign for hod approval",
    }

    status_keys = (
        "approverStatus",
        "ApproverStatus",
        "ApprovStatus",
        "inviteStatus",
        "status",
        "Status",
    )
    current_status = str(
        _first_existing_record_value(
            db_record,
            local_record,
            *status_keys,
        )
        or ""
    ).strip()
    current_status_lower = re.sub(r"\s+", " ", current_status.lower()).strip().replace("pendign", "pending")

    def _resolve_vendor_decision_role(base_role: str, mode_value: str) -> str:
        role_key = normalize_role(base_role)
        mode_key = str(mode_value or "").strip().lower()
        mode_stage = {
            "buyer": "user",
            "user": "user",
            "validator": "validator",
            "hod": "hod",
        }.get(mode_key, "")

        if role_key == "validator_user":
            if mode_stage in {"validator", "user"}:
                return mode_stage
            if current_status_lower in pending_user_statuses:
                return "user"
            return "validator"

        if role_key == "validator_hod":
            if mode_stage in {"validator", "hod"}:
                return mode_stage
            if current_status_lower in pending_hod_statuses:
                return "hod"
            return "validator"

        return role_key

    normalized_role = _resolve_vendor_decision_role(str(user.get("role") or ""), approval_mode)
    user_email = str(user.get("email") or "").strip().lower()
    user_email_raw = str(user.get("email") or "").strip()
    if normalized_role == "hod":
        allowed_hod_emails = {
            email
            for email in (
                _extract_record_hod_email(local_record),
                _extract_record_hod_email(db_record),
            )
            if email
        }
        if not user_email or not allowed_hod_emails or user_email not in allowed_hod_emails:
            raise HTTPException(status_code=403, detail="Access denied for this record")
    elif normalized_role not in {"admin", "validator"}:
        allowed_emails = _collect_record_access_emails(local_record, db_record)
        if not user_email or not allowed_emails or user_email not in allowed_emails:
            raise HTTPException(status_code=403, detail="Access denied for this record")

    def is_pending_for_role(role_key: str) -> bool:
        if not current_status_lower:
            return False

        normalized_role_key = str(role_key or "").strip().lower()
        if normalized_role_key == "validator":
            if current_status_lower in pending_validator_statuses:
                return True
            return "pending" in current_status_lower and "validator" in current_status_lower
        if normalized_role_key == "hod":
            if current_status_lower in pending_hod_statuses:
                return True
            return "pending" in current_status_lower and "hod" in current_status_lower

        if current_status_lower in pending_user_statuses:
            return True
        return (
            "pending" in current_status_lower
            and (
                "user" in current_status_lower
                or "buyer" in current_status_lower
                or "data validation" in current_status_lower
            )
        )

    def guard_repeated_action(role_key: str) -> None:
        if not current_status_lower:
            raise HTTPException(
                status_code=409,
                detail="Record current approval status is missing. Please refresh the record before submitting a decision.",
            )
        if is_pending_for_role(role_key):
            return
        if decision == "approve":
            role_label = str(role_key or "current").strip().lower() or "current"
            raise HTTPException(
                status_code=409,
                detail=f"Record is not pending for {role_label} approval. Current status: {current_status}",
            )
        raise HTTPException(
            status_code=409,
            detail=f"Decision already submitted for this record. Current status: {current_status}",
        )

    if normalized_role == "validator":
        guard_repeated_action("validator")
    elif normalized_role == "hod":
        guard_repeated_action("hod")
    else:
        guard_repeated_action("user")

    user_role = str(user.get("role") or "User").strip() or "User"
    decision_label = "Approve" if decision == "approve" else "Reject"
    updates: dict[str, Any] = {}

    if normalized_role == "validator":
        if decision == "approve":
            approver_status = "Pending For HOD Approval"
        else:
            approver_status = "Rejected by Validator"
            if reject_to.lower() not in {"vendor", "buyer"}:
                raise HTTPException(status_code=400, detail="rejectTo must be Vendor or Buyer")

        def parse_bool(value: Any) -> bool:
            if isinstance(value, bool):
                return value
            if value is None:
                return False
            return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}

        def parse_string_value(value: Any) -> str | None:
            if isinstance(value, (list, tuple, set)):
                normalized_list = [str(item).strip() for item in value if str(item).strip()]
                return normalized_list[0] if normalized_list else None
            if value is None:
                return None
            text = str(value).strip()
            if not text:
                return None
            try:
                decoded = json.loads(text)
            except Exception:
                decoded = None
            if isinstance(decoded, list):
                normalized_list = [str(item).strip() for item in decoded if str(item).strip()]
                return normalized_list[0] if normalized_list else None
            if decoded is not None and not isinstance(decoded, dict):
                normalized = str(decoded).strip()
                return normalized or None
            if text.startswith("[") and text.endswith("]"):
                first_value = next(
                    (
                        item.strip().strip("'").strip('"')
                        for item in text[1:-1].split(",")
                        if item.strip().strip("'").strip('"')
                    ),
                    "",
                )
                return first_value or None
            return text

        validations_payload = validations_raw if isinstance(validations_raw, dict) else {}
        gst_validate = parse_bool(validations_payload.get("gstValidate"))
        cin_validate = parse_bool(validations_payload.get("cinValidate"))
        pan_validate = parse_bool(validations_payload.get("panValidate"))
        msme_validate = parse_bool(validations_payload.get("msmeValidate"))
        pf_validate = parse_bool(validations_payload.get("pfValidate"))
        esi_validate = parse_bool(validations_payload.get("esiValidate"))
        address_validate = parse_bool(validations_payload.get("addressValidate"))
        bank_validate = parse_bool(validations_payload.get("bankValidate"))
        gst_validate_text = "True" if gst_validate else "False"
        cin_validate_text = "True" if cin_validate else "False"
        pan_validate_text = "True" if pan_validate else "False"
        msme_validate_text = "True" if msme_validate else "False"
        pf_validate_text = "True" if pf_validate else "False"
        esi_validate_text = "True" if esi_validate else "False"
        address_validate_text = "True" if address_validate else "False"
        bank_validate_text = "True" if bank_validate else "False"
        tax_payload = tax_details_raw if isinstance(tax_details_raw, dict) else {}
        withholding_tax_raw = tax_payload.get("withholdingTax")
        if withholding_tax_raw is None:
            withholding_tax_raw = body.get("withholdingTax")
        turnover_limit_raw = tax_payload.get("turnoverLimit")
        if turnover_limit_raw is None:
            turnover_limit_raw = body.get("turnoverLimit")
        withholding_tax_value = parse_string_value(withholding_tax_raw)
        turnover_limit_value = parse_string_value(turnover_limit_raw)
        updates.update(
            {
                "approverStatus": approver_status,
                "ApproverStatus": approver_status,
                "reviewDecision": decision_label,
                "ReviewDecision": decision_label,
                "reviewRemarks": remarks,
                "ReviewRemarks": remarks,
                "reviewedByRole": user_role,
                "ReviewedByRole": user_role,
                "remarks": remarks,
                "Remarks": remarks,
                "gstValidate": gst_validate_text,
                "cinValidate": cin_validate_text,
                "panValidate": pan_validate_text,
                "msmeValidate": msme_validate_text,
                "pfValidate": pf_validate_text,
                "esiValidate": esi_validate_text,
                "addressValidate": address_validate_text,
                "bankValidate": bank_validate_text,
                "GST Validate": gst_validate_text,
                "CIN Validate": cin_validate_text,
                "PAN Validate": pan_validate_text,
                "MSME Validate": msme_validate_text,
                "PF Validate": pf_validate_text,
                "ESI Validate": esi_validate_text,
                "Address Validate": address_validate_text,
                "Bank Validate": bank_validate_text,
                "validateBy": user_email_raw or user_email,
                "Validate By": user_email_raw or user_email,
                "withholdingTax": withholding_tax_value,
                "WithholdingTax": withholding_tax_value,
                "Withholding Tax": withholding_tax_value,
                "turnoverLimit": turnover_limit_value,
                "TurnoverLimit": turnover_limit_value,
                "Turnover Limit": turnover_limit_value,
            }
        )
        if reject_to:
            updates["rejectTo"] = reject_to.title()
            updates["Reject To"] = reject_to.title()
    elif normalized_role == "hod":
        if decision == "approve":
            approver_status = STATUS_PENDING_SAP_CODE_CREATION
        else:
            approver_status = "Rejected by HOD"
            if reject_to.lower() not in {"vendor", "buyer"}:
                raise HTTPException(status_code=400, detail="rejectTo must be Vendor or Buyer")

        updates.update(
            {
                "approverStatus": approver_status,
                "ApproverStatus": approver_status,
                "reviewDecision": decision_label,
                "ReviewDecision": decision_label,
                "reviewRemarks": remarks,
                "ReviewRemarks": remarks,
                "reviewedByRole": user_role,
                "ReviewedByRole": user_role,
                "remarks": remarks,
                "Remarks": remarks,
            }
        )
        if reject_to:
            updates["rejectTo"] = reject_to.title()
            updates["Reject To"] = reject_to.title()
    else:
        approver_status = (
            "Pendign For Validator Approval"
            if decision == "approve"
            else f"Rejected by {user_role}"
        )

        updates.update(
            {
                "approverStatus": approver_status,
                "ApproverStatus": approver_status,
                "reviewDecision": decision_label,
                "ReviewDecision": decision_label,
                "reviewRemarks": remarks,
                "ReviewRemarks": remarks,
                "reviewedByRole": user_role,
                "ReviewedByRole": user_role,
                "remarks": remarks,
                "Remarks": remarks,
            }
        )

        team_payload = team_raw if isinstance(team_raw, dict) else {}
        approve_required_team_fields = ("incoTerms", "paymentTerms", "vendorType", "rubaminApproverHod")
        if decision == "approve":
            missing_fields = []
            for field_name in approve_required_team_fields:
                field_value = empty_to_none(team_payload.get(field_name)) if isinstance(team_payload.get(field_name), str) else team_payload.get(field_name)
                if field_value is None or (isinstance(field_value, str) and field_value.strip() == ""):
                    missing_fields.append(field_name)
            if missing_fields:
                raise HTTPException(
                    status_code=400,
                    detail=f"Missing required fields for approve: {', '.join(missing_fields)}",
                )

        team_field_keys = {
            "incoTerms": ["IncoTerms", "Incoterms"],
            "paymentTerms": ["PaymentTerms"],
            "vendorType": ["VendorType"],
            "rubaminApproverHod": [
                "RubaminApproverHod",
                "Rubamin Approver HOD",
                "RubaminApprovalHod",
                "hodEmail",
                "HodEmail",
                "HODEmail",
            ],
        }
        for canonical_key, alias_keys in team_field_keys.items():
            if canonical_key not in team_payload:
                continue
            raw_value = team_payload.get(canonical_key)
            normalized_value = empty_to_none(raw_value) if isinstance(raw_value, str) else raw_value
            updates[canonical_key] = normalized_value
            for alias_key in alias_keys:
                updates[alias_key] = normalized_value

    columns_to_ensure = [
        "ApproverStatus",
        "ReviewDecision",
        "ReviewRemarks",
        "ReviewedByRole",
        "Remarks",
    ]
    if normalized_role == "validator":
        columns_to_ensure.extend(
            [
                "GST Validate",
                "CIN Validate",
                "PAN Validate",
                "MSME Validate",
                "PF Validate",
                "ESI Validate",
                "Address Validate",
                "Bank Validate",
                "Validate By",
                "WithholdingTax",
                "TurnoverLimit",
            ]
        )

    try:
        ensure_vendor_columns(columns_to_ensure)
    except Exception:
        pass

    db_updated = update_vendor_record_by_record_id(record_id, updates)
    if db_record and not db_updated:
        raise HTTPException(status_code=500, detail="Vendor record found in database but update failed")

    local_updated = update_local_vendor_record_by_record_id(record_id, updates)
    if not local_updated and not db_updated:
        raise HTTPException(status_code=400, detail="Unable to update vendor review decision")

    refreshed_local = find_local_vendor_record_by_record_id(record_id)
    refreshed_db = get_vendor_record_by_record_id(record_id)
    merged_record: dict[str, Any] = {}
    if refreshed_db:
        merged_record.update(refreshed_db)
    if refreshed_local:
        merged_record.update(refreshed_local)
    if not merged_record:
        merged_record = dict(updates)
        merged_record["recordId"] = record_id

    dashboard_url = _build_public_registration_url(request, "/web/vendor-dashboard")
    vendor_name = str(
        _first_existing_record_value(
            merged_record,
            None,
            "vendorName",
            "Vendorname",
            "VendorName",
            "name",
            "Name",
        )
        or ""
    )
    submitted_on = str(
        _first_existing_record_value(
            merged_record,
            None,
            "updatedAt",
            "UpdatedAt",
            "createdAt",
            "CreatedAt",
        )
        or now_iso()
    )
    update_vendor_form_url = _build_public_registration_url(request, "/web/update-vendor")
    vendor_email = (
        _extract_record_vendor_email(merged_record)
        or _extract_record_vendor_email(local_record)
        or _extract_record_vendor_email(db_record)
    )
    vendor_name_for_mail = (
        _extract_record_vendor_name(merged_record)
        or _extract_record_vendor_name(local_record)
        or _extract_record_vendor_name(db_record)
        or vendor_name
    )
    buyer_email = (
        _extract_record_buyer_email(merged_record)
        or _extract_record_buyer_email(local_record)
        or _extract_record_buyer_email(db_record)
    )
    buyer_name = (
        _extract_record_buyer_name(merged_record)
        or _extract_record_buyer_name(local_record)
        or _extract_record_buyer_name(db_record)
    )

    validator_email_sent_count = 0
    validator_email_error: str | None = None
    if (
        decision == "approve"
        and normalized_role not in {"admin", "validator", "hod"}
        and str(approver_status or "").strip().lower() in pending_validator_statuses
    ):
        validation_errors: list[str] = []
        for contact in _get_validator_contacts():
            contact_email = str(contact.get("email") or "").strip()
            if not contact_email:
                continue
            contact_name = str(contact.get("name") or "").strip() or _derive_name_from_email(contact_email)
            try:
                send_vendor_submission_approval_email(
                    contact_person_name=contact_name,
                    contact_person_email=contact_email,
                    record_id=record_id,
                    vendor_name=vendor_name,
                    submitted_on=submitted_on,
                    approver_status=approver_status,
                    dashboard_url=dashboard_url,
                    remarks=remarks,
                )
                validator_email_sent_count += 1
            except Exception as exc:
                validation_errors.append(f"{contact_email}: {exc}")

        if validation_errors:
            validator_email_error = "; ".join(validation_errors)

    hod_approval_email_sent = False
    hod_approval_email_error: str | None = None
    if (
        decision == "approve"
        and normalized_role == "validator"
        and str(approver_status or "").strip().lower() in pending_hod_statuses
    ):
        hod_email = (
            _extract_record_hod_email(merged_record)
            or _extract_record_hod_email(local_record)
            or _extract_record_hod_email(db_record)
        )
        hod_name = _derive_name_from_email(hod_email)
        if hod_email:
            try:
                send_vendor_submission_approval_email(
                    contact_person_name=hod_name,
                    contact_person_email=hod_email,
                    record_id=record_id,
                    vendor_name=vendor_name,
                    submitted_on=submitted_on,
                    approver_status=approver_status,
                    dashboard_url=dashboard_url,
                    remarks=remarks,
                )
                hod_approval_email_sent = True
            except Exception as exc:
                hod_approval_email_error = str(exc)
        else:
            hod_approval_email_error = "HOD email not configured on this record"

    rejection_email_sent = False
    rejection_email_error: str | None = None
    if decision == "reject":
        try:
            if normalized_role == "validator":
                if reject_to.lower() == "vendor":
                    if not vendor_email:
                        raise RuntimeError("Vendor email not found")
                    send_vendor_rejection_email(
                        recipient_name=vendor_name_for_mail,
                        recipient_email=vendor_email,
                        record_id=record_id,
                        remarks=remarks,
                        update_form_url=update_vendor_form_url,
                        rejected_by="validator",
                        cc_email=buyer_email or None,
                    )
                    rejection_email_sent = True
                elif reject_to.lower() == "buyer":
                    if not buyer_email:
                        raise RuntimeError("Buyer email not found")
                    send_vendor_rejection_email(
                        recipient_name=buyer_name,
                        recipient_email=buyer_email,
                        record_id=record_id,
                        remarks=remarks,
                        update_form_url=update_vendor_form_url,
                        rejected_by="validator",
                    )
                    rejection_email_sent = True
            elif normalized_role == "hod":
                if reject_to.lower() == "vendor":
                    if not vendor_email:
                        raise RuntimeError("Vendor email not found")
                    send_vendor_rejection_email(
                        recipient_name=vendor_name_for_mail,
                        recipient_email=vendor_email,
                        record_id=record_id,
                        remarks=remarks,
                        update_form_url=update_vendor_form_url,
                        rejected_by="HOD",
                        cc_email=buyer_email or None,
                    )
                    rejection_email_sent = True
                elif reject_to.lower() == "buyer":
                    if not buyer_email:
                        raise RuntimeError("Buyer email not found")
                    send_vendor_rejection_email(
                        recipient_name=buyer_name,
                        recipient_email=buyer_email,
                        record_id=record_id,
                        remarks=remarks,
                        update_form_url=update_vendor_form_url,
                        rejected_by="HOD",
                    )
                    rejection_email_sent = True
            else:
                if not vendor_email:
                    raise RuntimeError("Vendor email not found")
                cc_buyer = buyer_email or user_email_raw
                send_vendor_rejection_email(
                    recipient_name=vendor_name_for_mail,
                    recipient_email=vendor_email,
                    record_id=record_id,
                    remarks=remarks,
                    update_form_url=update_vendor_form_url,
                    rejected_by="buyer",
                    cc_email=cc_buyer or None,
                )
                rejection_email_sent = True
        except Exception as exc:
            rejection_email_error = str(exc)

    sap_sync_started = False
    sap_sync_queued_count = 0
    if decision == "approve" and normalized_role == "hod":
        sap_sync_started, sap_sync_queued_count = _queue_sap_sync_for_pending_status(
            background_tasks,
            record_id,
            approver_status,
        )

    return {
        "message": "Vendor review decision submitted successfully",
        "recordId": record_id,
        "approverStatus": approver_status,
        "validatorApprovalEmailSentCount": validator_email_sent_count,
        "validatorApprovalEmailError": validator_email_error,
        "hodApprovalEmailSent": hod_approval_email_sent,
        "hodApprovalEmailError": hod_approval_email_error,
        "rejectionEmailSent": rejection_email_sent,
        "rejectionEmailError": rejection_email_error,
        "sapTemplateUpdated": None,
        "sapTemplateError": None,
        "vendorUploadUpdated": None,
        "vendorUploadError": None,
        "vendorUploadSync": None,
        "sapSyncStarted": sap_sync_started,
        "sapSyncQueuedCount": sap_sync_queued_count,
        "vendor": merged_record,
    }


@app.post("/api/vendor/hod-bulk-decision")
def submit_hod_bulk_decision(
    body: dict[str, Any],
    request: Request,
    background_tasks: BackgroundTasks,
    user: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    if not user_has_any_role(user, "hod"):
        raise HTTPException(status_code=403, detail="Only HOD can use this action")

    raw_record_ids = body.get("recordIds")
    decision = str(empty_to_none(body.get("decision")) or "").strip().lower()
    remarks = str(empty_to_none(body.get("remarks")) or "").strip()
    reject_to = str(empty_to_none(body.get("rejectTo")) or "").strip()
    user_email = str(user.get("email") or "").strip().lower()
    user_role = str(user.get("role") or "HOD").strip() or "HOD"

    if not isinstance(raw_record_ids, list) or not raw_record_ids:
        raise HTTPException(status_code=400, detail="recordIds must be a non-empty list")
    if decision not in {"approve", "reject"}:
        raise HTTPException(status_code=400, detail="decision must be Approve or Reject")
    if decision == "reject" and reject_to.lower() not in {"vendor", "buyer"}:
        raise HTTPException(status_code=400, detail="rejectTo must be Vendor or Buyer")
    if not remarks:
        raise HTTPException(status_code=400, detail="remarks are required")
    if not user_email:
        raise HTTPException(status_code=400, detail="User email not found")

    normalized_record_ids: list[str] = []
    seen_record_ids: set[str] = set()
    for raw_id in raw_record_ids:
        record_id = _normalize_record_id(raw_id)
        if not record_id:
            continue
        if record_id in seen_record_ids:
            continue
        seen_record_ids.add(record_id)
        normalized_record_ids.append(record_id)

    if not normalized_record_ids:
        raise HTTPException(status_code=400, detail="No valid recordIds provided")

    pending_status_keys = {
        "pending for hod approval",
        "pendign for hod approval",
    }
    target_status = (
        STATUS_PENDING_SAP_CODE_CREATION
        if decision == "approve"
        else "Rejected by HOD"
    )
    decision_label = "Approve" if decision == "approve" else "Reject"
    update_vendor_form_url = _build_public_registration_url(request, "/web/update-vendor")

    updated_record_ids: list[str] = []
    failed_records: list[dict[str, str]] = []
    rejection_email_sent_count = 0
    rejection_email_errors: list[str] = []
    sap_sync_record_ids: list[str] = []

    for record_id in normalized_record_ids:
        local_record = find_local_vendor_record_by_record_id(record_id)
        db_record = get_vendor_record_by_record_id(record_id)
        if not local_record and not db_record:
            failed_records.append({"recordId": record_id, "reason": "Vendor record not found"})
            continue

        assigned_hod_email = (
            _extract_record_hod_email(local_record)
            or _extract_record_hod_email(db_record)
        )
        if not assigned_hod_email or assigned_hod_email != user_email:
            failed_records.append({"recordId": record_id, "reason": "Record is not assigned to current HOD"})
            continue

        status_value = str(
            _first_existing_record_value(
                db_record,
                local_record,
                "approverStatus",
                "ApproverStatus",
                "ApprovStatus",
                "inviteStatus",
                "status",
                "Status",
            )
            or ""
        ).strip()
        if status_value.lower() not in pending_status_keys:
            failed_records.append({"recordId": record_id, "reason": "Record is not pending for HOD approval"})
            continue

        updates = {
            "approverStatus": target_status,
            "ApproverStatus": target_status,
            "ApprovStatus": target_status,
            "reviewDecision": decision_label,
            "ReviewDecision": decision_label,
            "reviewRemarks": remarks,
            "ReviewRemarks": remarks,
            "reviewedByRole": user_role,
            "ReviewedByRole": user_role,
            "remarks": remarks,
            "Remarks": remarks,
        }
        if decision == "reject" and reject_to:
            updates["rejectTo"] = reject_to.title()
            updates["Reject To"] = reject_to.title()

        db_updated = update_vendor_record_by_record_id(record_id, updates)
        local_updated = update_local_vendor_record_by_record_id(record_id, updates)
        if not db_updated and not local_updated:
            failed_records.append({"recordId": record_id, "reason": "Unable to update record"})
            continue

        updated_record_ids.append(record_id)

        if decision == "approve" and _is_pending_sap_code_creation_status(target_status):
            sap_sync_record_ids.append(record_id)

        if decision == "reject":
            buyer_email = (
                _extract_record_buyer_email(local_record)
                or _extract_record_buyer_email(db_record)
            )
            buyer_name = (
                _extract_record_buyer_name(local_record)
                or _extract_record_buyer_name(db_record)
            )
            vendor_email = (
                _extract_record_vendor_email(local_record)
                or _extract_record_vendor_email(db_record)
            )
            vendor_name = (
                _extract_record_vendor_name(local_record)
                or _extract_record_vendor_name(db_record)
            )
            try:
                if reject_to.lower() == "vendor":
                    if not vendor_email:
                        raise RuntimeError("Vendor email not found")
                    send_vendor_rejection_email(
                        recipient_name=vendor_name,
                        recipient_email=vendor_email,
                        record_id=record_id,
                        remarks=remarks,
                        update_form_url=update_vendor_form_url,
                        rejected_by="HOD",
                        cc_email=buyer_email or None,
                    )
                else:
                    if not buyer_email:
                        raise RuntimeError("Buyer email not found")
                    send_vendor_rejection_email(
                        recipient_name=buyer_name,
                        recipient_email=buyer_email,
                        record_id=record_id,
                        remarks=remarks,
                        update_form_url=update_vendor_form_url,
                        rejected_by="HOD",
                    )
                rejection_email_sent_count += 1
            except Exception as exc:
                rejection_email_errors.append(f"{record_id}: {exc}")

    if not updated_record_ids:
        raise HTTPException(
            status_code=400,
            detail="Unable to update selected records for HOD decision",
        )

    sap_sync_started = False
    if (
        decision == "approve"
        and _is_pending_sap_code_creation_status(target_status)
        and sap_sync_record_ids
    ):
        background_tasks.add_task(
            _run_sap_sync_pipeline_record_by_record_background,
            list(sap_sync_record_ids),
        )
        sap_sync_started = True

    return {
        "message": "HOD decision submitted successfully",
        "decision": decision_label,
        "approverStatus": target_status,
        "updatedRecordIds": updated_record_ids,
        "failedRecords": failed_records,
        "rejectionEmailSentCount": rejection_email_sent_count,
        "rejectionEmailError": "; ".join(rejection_email_errors) if rejection_email_errors else None,
        "sapTemplateUpdatedCount": None,
        "sapTemplateError": None,
        "vendorUploadUpdated": None,
        "vendorUploadError": None,
        "vendorUploadSync": None,
        "sapSyncStarted": sap_sync_started,
        "sapSyncQueuedCount": len(sap_sync_record_ids),
    }


@app.post("/api/vendor/search-duplicates")
def search_duplicate_vendors(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    vendor_pan = str(empty_to_none(body.get("vendorPan")) or "").strip().upper()
    if not vendor_pan:
        raise HTTPException(status_code=400, detail="vendorPan is required")
    if not PAN_REGEX.fullmatch(vendor_pan):
        raise HTTPException(status_code=400, detail="Invalid PAN format")

    matches = get_vendor_duplicate_records_by_pan(vendor_pan)
    return {
        "vendorPan": vendor_pan,
        "matchCount": len(matches),
        "matches": matches,
    }


@app.post("/api/vendor/invite")
def invite_vendor(
    body: dict[str, Any],
    request: Request,
    user: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    vendor_name = str(empty_to_none(body.get("vendorName")) or "").strip()
    vendor_email = str(empty_to_none(body.get("vendorEmail")) or "").strip()
    vendor_pan = str(empty_to_none(body.get("vendorPan")) or "").strip().upper()
    postal_zip_code = str(empty_to_none(body.get("postalZipCode")) or "").strip()
    vendor_type_raw = str(empty_to_none(body.get("vendorType")) or "").strip().lower()
    hod_email = str(empty_to_none(body.get("hodEmail")) or "").strip()

    if not vendor_name or not vendor_email or not vendor_type_raw:
        raise HTTPException(
            status_code=400,
            detail="vendorName, vendorEmail and vendorType are required",
        )

    if "@" not in vendor_email or "." not in vendor_email.split("@", 1)[-1]:
        raise HTTPException(status_code=400, detail="Invalid vendorEmail format")

    if hod_email and ("@" not in hod_email or "." not in hod_email.split("@", 1)[-1]):
        raise HTTPException(status_code=400, detail="Invalid hodEmail format")

    vendor_type_map = {"domestic": "Domestic", "import": "Import"}
    vendor_type = vendor_type_map.get(vendor_type_raw)
    if not vendor_type:
        raise HTTPException(status_code=400, detail="vendorType must be Domestic or Import")
    if vendor_type == "Domestic":
        if not vendor_pan:
            raise HTTPException(status_code=400, detail="vendorPan is required for Domestic vendorType")
        if not PAN_REGEX.fullmatch(vendor_pan):
            raise HTTPException(status_code=400, detail="Invalid PAN format")
        postal_zip_code = ""
    else:
        vendor_pan = ""

    record_id = str(uuid.uuid4()).upper()
    created_at = now_iso()
    rubamin_contact = resolve_rubamin_contact_person_details(body, user)
    registration_path = (
        "/web/vendor-registration"
        if vendor_type == "Domestic"
        else "/web/import-vendor"
    )
    registration_link = _build_public_registration_url(request, registration_path)

    invite_record = {
        "recordId": record_id,
        "vendorName": vendor_name,
        "vendorEmail": vendor_email,
        "contactPersonEmail": vendor_email,
        "vendorPan": vendor_pan or None,
        "postalZipCode": postal_zip_code or None,
        "vendorType": vendor_type,
        "approverStatus": "Vendor Invited",
        "hodEmail": hod_email or None,
        "createdAt": created_at,
        "status": "Invited",
        "invitedById": user["id"],
        "invitedByEmail": user["email"],
        "rubaminContactPerson": rubamin_contact.get("name"),
        "rubaminContactPersonEmail": rubamin_contact.get("email"),
        "rubaminContactPersonDepartment": rubamin_contact.get("department"),
        "registrationPath": registration_path,
        "registrationLink": registration_link,
        "sqlRecordId": None,
    }
    invite_forms[record_id] = invite_record

    sql_payload = {
        "recordId": record_id,
        "vendorName": vendor_name,
        "Vendorname": vendor_name,
        "vendorEmail": vendor_email,
        "ContactEmailID": vendor_email,
        "contactPersonEmail": vendor_email,
        "vendorPan": vendor_pan or None,
        "PANNo": vendor_pan or None,
        "panNumber": vendor_pan or None,
        "postalZipCode": postal_zip_code or None,
        "PostalZipCode": postal_zip_code or None,
        "vendorType": vendor_type,
        "VendorType": vendor_type,
        "ApproverStatus": "Vendor Invited",
        "Vendor_Category": vendor_type,
        "Rubamin Contact Person": rubamin_contact.get("name"),
        "Rubamin Contact Person Email": rubamin_contact.get("email"),
        "Rubamin Contact Person Department": rubamin_contact.get("department"),
        "hodEmail": hod_email or None,
        "inviteStatus": "Invited",
        "createdAt": created_at,
        "updatedAt": created_at,
        "invitedById": user["id"],
        "invitedByEmail": user["email"],
    }

    try:
        sql_record_id = insert_vendor_master_data(sql_payload)
        if sql_record_id:
            invite_record["sqlRecordId"] = sql_record_id
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Invite saved in memory but SQL insert failed: {exc}") from exc

    email_sent = False
    email_error: str | None = None
    try:
        send_vendor_invite_email(
            vendor_name=vendor_name,
            vendor_email=vendor_email,
            record_id=record_id,
            registration_url=registration_link,
            cc_email=str(user.get("email") or "").strip(),
        )
        email_sent = True
        invite_record["inviteEmailSentAt"] = now_iso()
    except Exception as exc:
        email_error = str(exc)
        invite_record["inviteEmailError"] = email_error

    return {
        "message": (
            "Vendor invite saved and email sent successfully"
            if email_sent
            else "Vendor invite saved, but email sending failed"
        ),
        "recordId": record_id,
        "registrationPath": registration_path,
        "registrationLink": registration_link,
        "emailSent": email_sent,
        "emailError": email_error,
        "invite": invite_record,
    }


@app.post("/api/vendor/register")
async def register_vendor(
    request: Request,
    background_tasks: BackgroundTasks,
    user: dict[str, Any] | None = Depends(get_optional_current_user),
) -> dict[str, Any]:
    form_data = await request.form()

    body, pending_uploads = await _extract_form_body_and_uploads(form_data)
    upload_urls: dict[str, str] = {}
    upload_sql_payload: dict[str, Any] = {}

    raw_record_id = empty_to_none(body.get("recordId"))
    record_id = _normalize_record_id(raw_record_id)
    existing_local_record: dict[str, Any] | None = None
    existing_db_record: dict[str, Any] | None = None

    if user is None:
        # External (vendor) submission must always use an existing invited recordId
        existing_local_record, existing_db_record = _get_existing_vendor_record_or_raise(record_id)
    else:
        # Internal submission:
        # - If a recordId is provided (e.g. editing an existing record), reuse it
        # - If not provided, create a brand‑new recordId so the form does not depend on an invite
        if record_id:
            existing_local_record, existing_db_record = _get_existing_vendor_record_or_raise(record_id)
        else:
            record_id = str(uuid.uuid4())
    for field_name, upload in pending_uploads.items():
        normalized_filename = str(upload.get("filename") or "").strip()
        content = upload.get("content") or b""
        content_type = str(upload.get("content_type") or "application/octet-stream")
        if not normalized_filename:
            continue

        upload_urls[field_name] = _store_uploaded_document(
            record_id,
            field_name,
            normalized_filename,
            content,
            content_type,
        )
        upload_sql_payload[f"{field_name}_FileName"] = normalized_filename
        upload_sql_payload[f"{field_name}_ContentType"] = content_type
        upload_sql_payload[f"{field_name}_Data"] = content
    actor = user or {"id": None, "name": "", "email": "", "role": "Vendor"}

    if user is None:
        invited_by_id = to_int(
            _first_existing_record_value(
                existing_local_record,
                existing_db_record,
                "invitedById",
            )
        )
        if invited_by_id is not None:
            body["buyerId"] = str(invited_by_id)

    validate_business_identifier_fields(body)
    agree_information_accuracy = truthy_flag(body.get("agreeInformationAccuracy"))
    agree_code_of_conduct = truthy_flag(body.get("agreeCodeOfConduct"))
    if user is None and not agree_information_accuracy:
        raise HTTPException(status_code=400, detail="agreeInformationAccuracy must be accepted")
    if user is None and not agree_code_of_conduct:
        raise HTTPException(status_code=400, detail="agreeCodeOfConduct must be accepted")

    iso_certificates: list[Any] = []
    if body.get("isoCertificates"):
        try:
            parsed = json.loads(body["isoCertificates"])
            if isinstance(parsed, list):
                iso_certificates = parsed
        except json.JSONDecodeError:
            iso_certificates = []

    form_id = next(form_id_counter)
    created_at = now_iso()
    if user is None:
        rubamin_contact = _merge_rubamin_contact_with_existing_record(
            {
                "name": "",
                "email": "",
                "department": empty_to_none(body.get("department")),
            },
            existing_local_record or existing_db_record,
        )
    else:
        rubamin_contact = _merge_rubamin_contact_with_existing_record(
            resolve_rubamin_contact_person_details(body, actor),
            existing_local_record or existing_db_record,
        )
    initial_approval_stage = "user" if user is None else "validator"
    initial_approver_status = (
        "Pending For User Approval"
        if initial_approval_stage == "user"
        else "Pending For Validator Approval"
    )

    vendor = {
        "id": form_id,
        "recordId": record_id,
        "vendorId": actor.get("id"),
        "buyerId": to_int(body.get("buyerId") or body.get("hodId")),
        "hodId": to_int(body.get("hodId")),
        "approverStatus": initial_approver_status,
        "buyerStatus": "pending",
        "buyerRemarks": None,
        "bankingStatus": "pending",
        "bankingRemarks": None,
        "accountsStatus": "pending",
        "accountsRemarks": None,
        "hodStatus": "pending",
        "hodRemarks": None,
        "taxationStatus": "pending",
        "taxationRemarks": None,
        "isoCertificates": iso_certificates,
        "agreeCodeOfConduct": agree_code_of_conduct,
        "agreeInformationAccuracy": agree_information_accuracy,
        "rubaminContactPerson": rubamin_contact.get("name"),
        "rubaminContactPersonEmail": rubamin_contact.get("email"),
        "rubaminContactPersonDepartment": rubamin_contact.get("department"),
        "createdAt": created_at,
        "updatedAt": created_at,
    }

    for key, value in body.items():
        if key in {"isoCertificates", "agreeCodeOfConduct", "agreeInformationAccuracy"}:
            continue
        vendor[key] = empty_to_none(value)

    for key in DOCUMENT_FIELDS:
        vendor[key] = upload_urls.get(key)

    for key, value in upload_urls.items():
        vendor[key] = value

    vendor_forms[form_id] = vendor
    update_local_vendor_record_by_record_id(record_id, vendor)

    try:
        ensure_vendor_columns(
            [
                "agreeCodeOfConduct",
                CODE_OF_CONDUCT_DOCUMENT_FIELD,
                f"{CODE_OF_CONDUCT_DOCUMENT_FIELD}_FileName",
                f"{CODE_OF_CONDUCT_DOCUMENT_FIELD}_ContentType",
                f"{CODE_OF_CONDUCT_DOCUMENT_FIELD}_Data",
            ]
        )
        sql_vendor_payload = dict(vendor)
        sql_vendor_payload["agreeCodeOfConduct"] = "True" if agree_code_of_conduct else "False"
        sql_vendor_payload["agreeInformationAccuracy"] = "True" if agree_information_accuracy else "False"
        sql_vendor_payload["Rubamin Contact Person"] = rubamin_contact.get("name")
        sql_vendor_payload["Rubamin Contact Person Email"] = rubamin_contact.get("email")
        sql_vendor_payload["Rubamin Contact Person Department"] = rubamin_contact.get("department")
        for key, value in upload_sql_payload.items():
            sql_vendor_payload[key] = value
        persisted_record_id = _persist_vendor_record_by_record_id(
            record_id,
            sql_vendor_payload,
            existing_db_record=existing_db_record,
        )
        vendor["recordId"] = persisted_record_id
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Vendor saved in memory but SQL insert failed: {exc}")

    approval_email_sent = False
    approval_email_error: str | None = None
    approval_email_sent_count = 0
    try:
        approval_email_sent_count, approval_email_error = _send_stage_approval_notifications(
            stage=initial_approval_stage,
            record_id=str(vendor.get("recordId") or record_id),
            vendor_name=str(vendor.get("vendorName") or ""),
            submitted_on=created_at,
            approver_status=str(vendor.get("approverStatus") or ""),
            dashboard_url=_build_public_registration_url(request, "/web/vendor-dashboard"),
            merged_record=vendor,
            local_record=existing_local_record,
            db_record=existing_db_record,
            preferred_contact_name=str(rubamin_contact.get("name") or ""),
            preferred_contact_email=str(rubamin_contact.get("email") or ""),
        )
        approval_email_sent = approval_email_sent_count > 0
        if approval_email_sent:
            vendor["approvalEmailSentAt"] = now_iso()
        if approval_email_error:
            vendor["approvalEmailError"] = approval_email_error
    except Exception as exc:
        approval_email_error = str(exc)
        vendor["approvalEmailError"] = approval_email_error

    background_tasks.add_task(
        _run_vendor_post_submit_enrichment,
        str(vendor.get("recordId") or record_id),
    )

    return {
        "message": "Vendor registered successfully",
        "recordId": vendor.get("recordId") or record_id,
        "approvalEmailSent": approval_email_sent,
        "approvalEmailSentCount": approval_email_sent_count,
        "approvalEmailError": approval_email_error,
        "vendor": hydrate_form(vendor),
    }


@app.post("/api/vendor/import-register")
async def register_import_vendor(
    request: Request,
    user: dict[str, Any] = Depends(get_current_user_from_token_or_cookie),
) -> dict[str, Any]:
    form_data = await request.form()

    body, pending_uploads = await _extract_form_body_and_uploads(form_data)
    upload_urls: dict[str, str] = {}
    upload_sql_payload: dict[str, Any] = {}

    record_id = str(uuid.uuid4()).upper()
    existing_local_record: dict[str, Any] | None = None
    existing_db_record: dict[str, Any] | None = None
    for field_name, upload in pending_uploads.items():
        normalized_filename = str(upload.get("filename") or "").strip()
        content = upload.get("content") or b""
        content_type = str(upload.get("content_type") or "application/octet-stream")
        if not normalized_filename:
            continue

        upload_urls[field_name] = _store_uploaded_document(
            record_id,
            field_name,
            normalized_filename,
            content,
            content_type,
        )
        upload_sql_payload[f"{field_name}_FileName"] = normalized_filename
        upload_sql_payload[f"{field_name}_ContentType"] = content_type
        upload_sql_payload[f"{field_name}_Data"] = content
    actor = user

    selected_sections_raw = str(body.get("selectedSections") or "").strip()
    selected_sections: set[str] = set()
    if selected_sections_raw:
        try:
            parsed_selected_sections = json.loads(selected_sections_raw)
            if isinstance(parsed_selected_sections, list):
                for item in parsed_selected_sections:
                    section_name = str(item or "").strip().lower()
                    if section_name:
                        selected_sections.add(section_name)
        except Exception:
            selected_sections = set()

    def section_selected(*keywords: str) -> bool:
        if not selected_sections:
            return True
        normalized_keywords = [str(keyword or "").strip().lower() for keyword in keywords if str(keyword or "").strip()]
        if not normalized_keywords:
            return True
        for section_name in selected_sections:
            if any(keyword in section_name for keyword in normalized_keywords):
                return True
        return False

    field_section_keywords: dict[str, tuple[str, ...]] = {
        "companyName": ("vendor", "basic"),
        "vendorName": ("vendor", "basic"),
        "vendorCategory": ("vendor", "basic"),
        "vendorType": ("commercial",),
        "rubaminApproverHod": ("commercial",),
        "addressLane1": ("address",),
        "addressLane2": ("address",),
        "country": ("address",),
        "state": ("address",),
        "district": ("address",),
        "pincode": ("address",),
        "currency": ("commercial",),
        "paymentTerms": ("commercial",),
        "incoTerms": ("commercial",),
        "serviceProvideFor": ("commercial",),
        "materialDealingsIn": ("commercial",),
        "form10fStatus": ("commercial",),
        "contactPersonName": ("contact",),
        "contactPersonEmail": ("contact",),
        "contactPersonMobile": ("contact",),
    }

    missing_fields = [
        field_name
        for field_name in IMPORT_REQUIRED_FIELDS
        if section_selected(*field_section_keywords.get(field_name, tuple()))
        if str(empty_to_none(body.get(field_name)) or "").strip() == ""
    ]
    if missing_fields:
        raise HTTPException(
            status_code=400,
            detail=f"Missing required fields: {', '.join(missing_fields)}",
        )

    contact_email = str(empty_to_none(body.get("contactPersonEmail")) or "").strip()
    if (section_selected("contact") or contact_email) and (
        "@" not in contact_email or "." not in contact_email.split("@", 1)[-1]
    ):
        raise HTTPException(status_code=400, detail="Invalid contactPersonEmail format")
    alternative_email = str(empty_to_none(body.get("alternativePersonEmail")) or "").strip()
    if alternative_email and ("@" not in alternative_email or "." not in alternative_email.split("@", 1)[-1]):
        raise HTTPException(status_code=400, detail="Invalid alternativePersonEmail format")

    form10f_status = str(empty_to_none(body.get("form10fStatus")) or "").strip().lower()
    commercial_selected = section_selected("commercial")
    if commercial_selected or form10f_status:
        if form10f_status not in {"yes", "no"}:
            raise HTTPException(status_code=400, detail="form10fStatus must be Yes or No")
        if form10f_status == "yes" and not upload_urls.get("form10fDocument"):
            raise HTTPException(status_code=400, detail="Form 10F document is required when Form 10F Status is Yes")
        if form10f_status == "no" and not upload_urls.get("trcDocument"):
            raise HTTPException(status_code=400, detail="TRC Document is required when Form 10F Status is No")

    banking_selected = section_selected("bank")
    bank_payment_method = truthy_flag(body.get("bankPaymentMethod")) if banking_selected else False
    if banking_selected and bank_payment_method:
        bank_required_fields = (
            "beneficiaryName",
            "bankName",
            "branchName",
            "bankAccountNumber",
            "swiftCode",
            "iban",
        )
        missing_bank_fields = [
            field_name
            for field_name in bank_required_fields
            if str(empty_to_none(body.get(field_name)) or "").strip() == ""
        ]
        if missing_bank_fields:
            raise HTTPException(
                status_code=400,
                detail=f"Missing required bank fields: {', '.join(missing_bank_fields)}",
            )
        if not upload_urls.get("bankDocument"):
            raise HTTPException(
                status_code=400,
                detail="Bank Document is required when Payment Method RTGS is enabled",
            )

    agree_information_accuracy = truthy_flag(body.get("agreeInformationAccuracy"))
    if not agree_information_accuracy:
        raise HTTPException(status_code=400, detail="agreeInformationAccuracy must be accepted")

    created_at = now_iso()
    rubamin_contact = _merge_rubamin_contact_with_existing_record(
        resolve_rubamin_contact_person_details(body, actor),
        existing_local_record or existing_db_record,
    )

    vendor_type = str(empty_to_none(body.get("vendorType")) or "Vendors-Foreign").strip() or "Vendors-Foreign"
    initial_approval_stage = "validator"
    approver_status = "Pending For Validator Approval"

    db_payload: dict[str, Any] = {
        "RecordID": record_id,
        "ApproverStatus": approver_status,
        "VendorType": vendor_type,
        "Rubamin Contact Person": rubamin_contact.get("name"),
        "Rubamin Contact Person Email": rubamin_contact.get("email"),
        "Rubamin Contact Person Department": rubamin_contact.get("department"),
        "createdAt": created_at,
        "updatedAt": created_at,
    }

    for form_key, column_name in IMPORT_FORM_FIELD_TO_COLUMN.items():
        if form_key == "vendorType":
            value = db_payload["VendorType"]
        elif form_key == "form10fStatus":
            value = "Yes" if form10f_status == "yes" else "No"
        elif form_key == "bankPaymentMethod":
            value = "True" if bank_payment_method else "False"
        elif form_key == "agreeInformationAccuracy":
            value = "True" if agree_information_accuracy else "False"
        else:
            value = empty_to_none(body.get(form_key))
        db_payload[column_name] = value

    hod_value = empty_to_none(body.get("rubaminApproverHod"))
    if hod_value is not None:
        db_payload["RubaminApprovalHod"] = hod_value
        db_payload["Rubamin Approver HOD"] = hod_value
        db_payload["Rubamin Approval HOD"] = hod_value
        db_payload["hodEmail"] = hod_value
        db_payload["HodEmail"] = hod_value
        db_payload["HODEmail"] = hod_value

    for form_key, column_name in IMPORT_FORM_UPLOAD_TO_COLUMN.items():
        db_payload[column_name] = upload_urls.get(form_key)
    for key, value in upload_sql_payload.items():
        db_payload[key] = value

    memory_form = {
        "id": next(form_id_counter),
        "recordId": record_id,
        "vendorType": vendor_type,
        "approverStatus": approver_status,
        "rubaminContactPerson": rubamin_contact.get("name"),
        "rubaminContactPersonEmail": rubamin_contact.get("email"),
        "rubaminContactPersonDepartment": rubamin_contact.get("department"),
        "createdAt": created_at,
        "updatedAt": created_at,
    }
    for form_key in IMPORT_FORM_FIELD_TO_COLUMN.keys():
        if form_key == "vendorType":
            memory_form[form_key] = vendor_type
        elif form_key == "form10fStatus":
            memory_form[form_key] = "Yes" if form10f_status == "yes" else "No"
        elif form_key == "bankPaymentMethod":
            memory_form[form_key] = bank_payment_method
        elif form_key == "agreeInformationAccuracy":
            memory_form[form_key] = agree_information_accuracy
        else:
            memory_form[form_key] = empty_to_none(body.get(form_key))
    if hod_value is not None:
        memory_form["RubaminApproverHod"] = hod_value
        memory_form["RubaminApprovalHod"] = hod_value
        memory_form["Rubamin Approver HOD"] = hod_value
        memory_form["hodEmail"] = hod_value
        memory_form["HodEmail"] = hod_value
        memory_form["HODEmail"] = hod_value
    for form_key in IMPORT_FORM_UPLOAD_TO_COLUMN.keys():
        memory_form[form_key] = upload_urls.get(form_key)
    vendor_forms[memory_form["id"]] = memory_form
    update_local_vendor_record_by_record_id(record_id, memory_form)

    try:
        persisted_record_id = _persist_vendor_record_by_record_id(
            record_id,
            db_payload,
            existing_db_record=existing_db_record,
        )
        memory_form["recordId"] = persisted_record_id
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Import vendor saved in memory but SQL insert failed: {exc}",
        ) from exc

    approval_email_sent_count = 0
    approval_email_error: str | None = None
    try:
        approval_email_sent_count, approval_email_error = _send_stage_approval_notifications(
            stage=initial_approval_stage,
            record_id=str(memory_form.get("recordId") or record_id),
            vendor_name=str(memory_form.get("vendorName") or ""),
            submitted_on=created_at,
            approver_status=str(memory_form.get("approverStatus") or ""),
            dashboard_url=_build_public_registration_url(request, "/web/vendor-dashboard"),
            merged_record=memory_form,
            local_record=existing_local_record,
            db_record=existing_db_record,
            preferred_contact_name=str(rubamin_contact.get("name") or ""),
            preferred_contact_email=str(rubamin_contact.get("email") or ""),
        )
        if approval_email_sent_count > 0:
            memory_form["approvalEmailSentAt"] = now_iso()
        if approval_email_error:
            memory_form["approvalEmailError"] = approval_email_error
    except Exception as exc:
        approval_email_error = str(exc)
        memory_form["approvalEmailError"] = approval_email_error

    return {
        "message": "Import vendor form submitted successfully",
        "recordId": memory_form.get("recordId") or record_id,
        "approvalEmailSent": approval_email_sent_count > 0,
        "approvalEmailSentCount": approval_email_sent_count,
        "approvalEmailError": approval_email_error,
        "vendor": memory_form,
    }


@app.get("/api/buyer/get-hod-emails")
def get_hod_emails(_: dict[str, Any] = Depends(require_role("Buyer"))) -> list[dict[str, Any]]:
    result = []
    for user in users.values():
        if user.get("role") == "Hod":
            result.append({"id": user["id"], "name": user["name"], "email": user["email"]})
    return result


def find_form_for_update(form_id: Any) -> dict[str, Any] | None:
    fid = to_int(form_id)
    if fid is None:
        return None
    return vendor_forms.get(fid)


def update_form(form: dict[str, Any], updates: dict[str, Any]) -> dict[str, Any]:
    form.update(updates)
    form["updatedAt"] = now_iso()
    return form


@app.get("/api/buyer/get-pending-vendor-form")
def buyer_pending(user: dict[str, Any] = Depends(require_role("Buyer"))) -> list[dict[str, Any]]:
    return sorted_forms({"buyerId": user["id"], "buyerStatus": "pending"})


@app.post("/api/buyer/submit-status-remarks-vendor-form")
def buyer_submit(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(require_role("Buyer")),
) -> dict[str, Any]:
    form_id = body.get("formId")
    buyer_status = body.get("buyerStatus")
    buyer_remarks = body.get("buyerRemarks")
    hod_id = body.get("hodId")

    if not form_id or not buyer_status or not buyer_remarks:
        raise HTTPException(
            status_code=400,
            detail={
                "error": "Missing required fields",
                "required": ["formId", "buyerStatus", "buyerRemarks"],
            },
        )

    if buyer_status == "approved" and not hod_id:
        raise HTTPException(
            status_code=400,
            detail={
                "error": "hodId is required when status is approved",
                "required": ["formId", "buyerStatus", "buyerRemarks", "hodId"],
            },
        )

    form = find_form_for_update(form_id)
    if not form or to_int(form.get("buyerId")) != user["id"]:
        raise HTTPException(status_code=404, detail="Form not found or not authorized")

    updates = {"buyerStatus": buyer_status, "buyerRemarks": buyer_remarks}
    if buyer_status == "approved":
        updates["hodId"] = to_int(hod_id)

    updated = update_form(form, updates)
    return {"message": "Form updated successfully", "form": hydrate_form(updated)}


@app.get("/api/buyer/get-vendor-forms")
def buyer_forms(user: dict[str, Any] = Depends(require_role("Buyer"))) -> list[dict[str, Any]]:
    return sorted_forms({"buyerId": user["id"]})


@app.get("/api/banking/get-pending-vendor-form")
def banking_pending(_: dict[str, Any] = Depends(require_role("Banking"))) -> list[dict[str, Any]]:
    return sorted_forms({"buyerStatus": "approved", "bankingStatus": "pending"})


@app.post("/api/banking/submit-status-remarks-vendor-form")
def banking_submit(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(require_role("Banking")),
) -> dict[str, Any]:
    form_id = body.get("formId")
    status = body.get("bankingStatus")
    remarks = body.get("bankingRemarks")

    if not form_id or not status or not remarks:
        raise HTTPException(
            status_code=400,
            detail={
                "error": "Missing required fields",
                "required": ["formId", "bankingStatus", "bankingRemarks"],
            },
        )

    form = find_form_for_update(form_id)
    if not form or form.get("bankingStatus") != "pending":
        raise HTTPException(status_code=404, detail="Form not found or not authorized")

    updated = update_form(
        form,
        {
            "bankingStatus": status,
            "bankingRemarks": remarks,
            "bankingReponserId": user["id"],
        },
    )
    return {"message": "Form updated successfully", "form": hydrate_form(updated)}


@app.get("/api/banking/get-vendor-forms")
def banking_forms(_: dict[str, Any] = Depends(require_role("Banking"))) -> list[dict[str, Any]]:
    return sorted_forms({"buyerStatus": "approved"})


@app.get("/api/accounts/get-pending-vendor-form")
def accounts_pending(_: dict[str, Any] = Depends(require_role("Accounts"))) -> list[dict[str, Any]]:
    return sorted_forms(
        {
            "buyerStatus": "approved",
            "bankingStatus": "approved",
            "accountsStatus": "pending",
        }
    )


@app.post("/api/accounts/submit-status-remarks-vendor-form")
def accounts_submit(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(require_role("Accounts")),
) -> dict[str, Any]:
    form_id = body.get("formId")
    status = body.get("accountsStatus")
    remarks = body.get("accountsRemarks")

    if not form_id or not status or not remarks:
        raise HTTPException(
            status_code=400,
            detail={
                "error": "Missing required fields",
                "required": ["formId", "accountsStatus", "accountsRemarks"],
            },
        )

    form = find_form_for_update(form_id)
    if not form or form.get("accountsStatus") != "pending":
        raise HTTPException(status_code=404, detail="Form not found or not authorized")

    updated = update_form(
        form,
        {
            "accountsStatus": status,
            "accountsRemarks": remarks,
            "accountsReponserId": user["id"],
        },
    )
    return {"message": "Form updated successfully", "form": hydrate_form(updated)}


@app.get("/api/accounts/get-vendor-forms")
def accounts_forms(_: dict[str, Any] = Depends(require_role("Accounts"))) -> list[dict[str, Any]]:
    return sorted_forms({"buyerStatus": "approved", "bankingStatus": "approved"})


@app.get("/api/hod/get-pending-vendor-form")
def hod_pending(_: dict[str, Any] = Depends(require_role("Hod"))) -> list[dict[str, Any]]:
    return sorted_forms(
        {
            "buyerStatus": "approved",
            "bankingStatus": "approved",
            "accountsStatus": "approved",
            "hodStatus": "pending",
        }
    )


@app.post("/api/hod/submit-status-remarks-vendor-form")
def hod_submit(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(require_role("Hod")),
) -> dict[str, Any]:
    form_id = body.get("formId")
    status = body.get("hodStatus")
    remarks = body.get("hodRemarks")

    if not form_id or not status or not remarks:
        raise HTTPException(
            status_code=400,
            detail={
                "error": "Missing required fields",
                "required": ["formId", "hodStatus", "hodRemarks"],
            },
        )

    form = find_form_for_update(form_id)
    if not form or form.get("hodStatus") != "pending":
        raise HTTPException(status_code=404, detail="Form not found or not authorized")

    updated = update_form(
        form,
        {
            "hodStatus": status,
            "hodRemarks": remarks,
            "hodReponserId": user["id"],
        },
    )
    return {"message": "Form updated successfully", "form": hydrate_form(updated)}


@app.get("/api/hod/get-vendor-forms")
def hod_forms(_: dict[str, Any] = Depends(require_role("Hod"))) -> list[dict[str, Any]]:
    return sorted_forms(
        {
            "buyerStatus": "approved",
            "bankingStatus": "approved",
            "accountsStatus": "approved",
        }
    )


@app.get("/api/taxation/get-pending-vendor-form")
def taxation_pending(_: dict[str, Any] = Depends(require_role("Taxation"))) -> list[dict[str, Any]]:
    return sorted_forms(
        {
            "buyerStatus": "approved",
            "bankingStatus": "approved",
            "accountsStatus": "approved",
            "hodStatus": "approved",
            "taxationStatus": "pending",
        }
    )


@app.post("/api/taxation/submit-status-remarks-vendor-form")
def taxation_submit(
    body: dict[str, Any],
    user: dict[str, Any] = Depends(require_role("Taxation")),
) -> dict[str, Any]:
    form_id = body.get("formId")
    status = body.get("taxationStatus")
    remarks = body.get("taxationRemarks")

    if not form_id or not status or not remarks:
        raise HTTPException(
            status_code=400,
            detail={
                "error": "Missing required fields",
                "required": ["formId", "taxationStatus", "taxationRemarks"],
            },
        )

    form = find_form_for_update(form_id)
    if not form or form.get("taxationStatus") != "pending":
        raise HTTPException(status_code=404, detail="Form not found or not authorized")

    updated = update_form(
        form,
        {
            "taxationStatus": status,
            "taxationRemarks": remarks,
            "taxationReponserId": user["id"],
        },
    )
    return {"message": "Form updated successfully", "form": hydrate_form(updated)}


@app.get("/api/taxation/get-vendor-forms")
def taxation_forms(_: dict[str, Any] = Depends(require_role("Taxation"))) -> list[dict[str, Any]]:
    return sorted_forms(
        {
            "buyerStatus": "approved",
            "bankingStatus": "approved",
            "accountsStatus": "approved",
            "hodStatus": "approved",
        }
    )


def get_default_approval_sequence() -> dict[str, str]:
    return {
        "one": "Buyer",
        "two": "Banking",
        "three": "Accounts",
        "four": "Hod",
        "five": "Taxation",
    }


@app.post("/api/approval-sequence/initialize-default")
def initialize_default_approval_sequence() -> dict[str, Any]:
    global approval_sequence

    defaults = get_default_approval_sequence()
    if approval_sequence:
        approval_sequence.update(defaults)
        approval_sequence["updatedAt"] = now_iso()
        return {
            "message": "Default approval sequence initialized successfully (updated existing)",
            "approvalSequence": approval_sequence,
        }

    approval_sequence = {
        "id": next(approval_id_counter),
        **defaults,
        "createdAt": now_iso(),
        "updatedAt": now_iso(),
    }
    return {
        "message": "Default approval sequence initialized successfully (created new)",
        "approvalSequence": approval_sequence,
    }


@app.put("/api/approval-sequence/")
def update_approval_sequence(body: dict[str, Any]) -> dict[str, Any]:
    global approval_sequence

    if not approval_sequence:
        raise HTTPException(status_code=404, detail="Approval sequence not found. Please create one first.")

    update_data = {}
    for field in ["one", "two", "three", "four", "five"]:
        if field in body and body[field] is not None:
            update_data[field] = body[field]

    if not update_data:
        raise HTTPException(status_code=400, detail="At least one field must be provided for update")

    approval_sequence.update(update_data)
    approval_sequence["updatedAt"] = now_iso()

    return {
        "message": "Approval sequence updated successfully",
        "approvalSequence": approval_sequence,
    }


@app.get("/api/approval-sequence/")
def get_approval_sequence() -> dict[str, Any]:
    if not approval_sequence:
        raise HTTPException(status_code=404, detail="Approval sequence not found")
    return approval_sequence


# Register customer-only API endpoints (kept separate from the vendor backend code).
try:
    from .customer_backend import register_customer_routes as _register_customer_routes

    _register_customer_routes(app)
except Exception:
    # Avoid hard crash if customer routes are temporarily unavailable during development.
    pass
