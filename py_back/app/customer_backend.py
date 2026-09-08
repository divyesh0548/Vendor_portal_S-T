from __future__ import annotations

import json
import mimetypes
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from fastapi import Depends, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.datastructures import UploadFile as StarletteUploadFile

from . import customer_table_backend as customer_db


def register_customer_routes(app) -> None:
    """
    Attach customer-only API endpoints to the main FastAPI `app`.

    Imports from `app.main` are done lazily at registration time to avoid circular imports.
    """

    from . import main as m

    # Remove obsolete customer columns if they still exist.
    try:
        customer_db.remove_obsolete_customer_gst_columns()
    except Exception:
        pass

    # Collapse duplicate/legacy customer columns (including status aliases).
    try:
        customer_db.cleanup_customer_duplicate_columns()
    except Exception:
        pass

    def pick_first(record: dict[str, Any] | None, *keys: str) -> str:
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

    def get_customer_status_value(record: dict[str, Any] | None) -> str:
        if not isinstance(record, dict):
            return ""
        for key in (
            "approverStatus",
            "ApproverStatus",
            "ApprovStatus",
            "inviteStatus",
            "InviteStatus",
            "status",
            "Status",
        ):
            value = record.get(key)
            if value is None:
                continue
            text = str(value).strip()
            if text:
                return text
        return ""

    CUSTOMER_RECORD_INDICATOR_KEYS = (
        "CustomerName",
        "customerName",
        "CompanyDealing",
        "CompanyDealingWith",
        "CompanyStatus",
        "FirmType",
        "ManufacturingType",
        "TypeofEntity",
        "CustomerGroup",
        "CustomerType",
        "PANNumber",
        "GSTIN",
        "TradeNameGST",
        "FurnishBy",
        "AP_FurnishBy1",
        "Datepicker",
        "HODEmail",
        "SalesOrganization",
        "DistributionChannel",
        "AccountAssignment",
        "ScrapSales",
        "FileUploader",
        "FileUploader1",
        "FileUploader2",
        "FileUploader3",
        "FileUploader4",
        "FileUploader5",
    )

    def is_customer_record(record: dict[str, Any] | None) -> bool:
        if not isinstance(record, dict):
            return False

        for key in CUSTOMER_RECORD_INDICATOR_KEYS:
            if key in record:
                return True

        type_value = pick_first(record, "CustomerType", "vendorType", "VendorType", "TypeofEntity")
        return "customer" in str(type_value or "").strip().lower()

    def iter_customer_invite_records():
        for record in m.invite_forms.values():
            if is_customer_record(record):
                yield record

    def iter_customer_local_form_records():
        for record in m.vendor_forms.values():
            if is_customer_record(record):
                yield record

    def map_customer_summary(record: dict[str, Any]) -> dict[str, Any]:
        status_value = get_customer_status_value(record)
        return {
            "record": pick_first(record, "recordId", "RecordID", "record", "Record", "id", "Id", "ID"),
            "company": pick_first(
                record,
                "companyName",
                "CompanyName",
                "CompanyDealing",
                "CompanyDealingWith",
                "company",
                "CustomerGroup",
                "TopCustomer",
                "topCustomer",
            ),
            "category": pick_first(
                record,
                "FirmType",
                "firmType",
                "CustomerGroup",
                "customerGroup",
                "vendorCategory",
                "VendorCategory",
                "category",
            ),
            "approverStatus": str(status_value or ""),
            "name": pick_first(
                record,
                "CustomerName",
                "customerName",
                "vendorName",
                "VendorName",
                "TopCustomer",
                "topCustomer",
                "name",
            ),
            "type": pick_first(
                record,
                "CustomerType",
                "vendorType",
                "VendorType",
                "TypeofEntity",
                "IndustryType",
                "DistributionChannel",
                "type",
            ),
            "rubaminContactPerson": pick_first(
                record,
                "rubaminContactPerson",
                "Rubamin Contact Person",
            ),
            "rubaminContactPersonEmail": pick_first(
                record,
                "rubaminContactPersonEmail",
                "Rubamin Contact Person Email",
                "rubaminContactPerson",
                "Rubamin Contact Person",
            ),
            "createdAt": pick_first(
                record, "createdAt", "CreatedAt", "created_at", "updatedAt", "UpdatedAt", "updated_at"
            ),
        }

    def is_customer_user_assigned_to_rubamin_contact(
        record: dict[str, Any] | None,
        target_email: str,
    ) -> bool:
        if not isinstance(record, dict):
            return False
        normalized_target = str(target_email or "").strip().lower()
        if not normalized_target:
            return False

        for key in (
            "rubaminContactPersonEmail",
            "Rubamin Contact Person Email",
            "rubaminContactPerson",
            "Rubamin Contact Person",
        ):
            for candidate in m._extract_email_candidates(record.get(key)):
                if candidate == normalized_target:
                    return True
        # Backward-compatible fallback for legacy records that may still rely on
        # older assignee columns (buyer/contact/invitedBy aliases).
        return m._record_is_assigned_to_email(record, normalized_target)

    def load_customer_sql_records() -> list[dict[str, Any]]:
        try:
            return customer_db.get_all_customer_records()
        except Exception:
            return []

    def get_customer_records(record_id: str) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
        local_record = m.find_local_vendor_record_by_record_id(record_id)
        if local_record and not is_customer_record(local_record):
            local_record = None
        sql_record = customer_db.get_customer_record_by_record_id(record_id)
        return local_record, sql_record

    def merge_customer_records(
        record_id: str,
        local_record: dict[str, Any] | None,
        sql_record: dict[str, Any] | None,
    ) -> dict[str, Any]:
        merged_record: dict[str, Any] = {}
        if sql_record:
            merged_record.update(sql_record)
        if local_record:
            merged_record.update(local_record)
        sql_status = str(
            m._first_existing_record_value(
                sql_record,
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
        if sql_status:
            merged_record["approverStatus"] = sql_status
            merged_record["ApproverStatus"] = sql_status
            merged_record["ApprovStatus"] = sql_status
        merged_record["recordId"] = record_id
        return merged_record

    def authorize_customer_record_access(
        record_id: str,
        local_record: dict[str, Any] | None,
        sql_record: dict[str, Any] | None,
        user: dict[str, Any] | None,
    ) -> None:
        # Keep optional auth behavior for compatibility with existing flows.
        if user is None:
            return

        normalized_role = m.normalize_role(user.get("role"))
        user_email = str(user.get("email") or "").strip().lower()
        bot_name = m.normalize_bot_name(user.get("botName") or "Vendor Bot")

        if m.user_has_any_role(user, "admin", "validator"):
            return

        if m.user_has_any_role(user, "hod"):
            allowed_hod_emails = {
                email
                for email in (
                    m._extract_record_hod_email(local_record),
                    m._extract_record_hod_email(sql_record),
                )
                if email
            }
            if not user_email or not allowed_hod_emails or user_email not in allowed_hod_emails:
                raise HTTPException(status_code=403, detail="Access denied for this record")
            return

        # Compatibility guard:
        # Some customer flows land with `User` role token while navigating Validator/HOD queues.
        # Allow read-only view/document access for customer-enabled bots.
        if m.user_has_any_role(user, "user", "buyer") and bot_name in {"Both", "Customer Bot"}:
            return

        allowed_emails = m._collect_record_access_emails(local_record, sql_record)
        if user_email and user_email in allowed_emails:
            return

        if user_email:
            for row in customer_db.get_customer_records_by_assignee_email(user_email):
                row_record = m._normalize_record_id(row.get("record"))
                if row_record and row_record.lower() == record_id.lower():
                    return

        raise HTTPException(status_code=403, detail="Access denied for this record")

    def get_customer_display_name(*records: dict[str, Any] | None) -> str:
        for record in records:
            if not isinstance(record, dict):
                continue
            name_value = pick_first(
                record,
                "CustomerName",
                "vendorName",
                "VendorName",
                "Vendorname",
                "TopCustomer",
                "name",
                "Name",
            )
            if name_value:
                return name_value
        return ""

    def parse_bool(value: Any) -> bool:
        if isinstance(value, bool):
            return value
        if value is None:
            return False
        return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}

    STATUS_INVITE_CUSTOMER = "Invite Customer"
    STATUS_PENDING_USER_APPROVAL = "Pending for User Approval"
    STATUS_PENDING_VALIDATOR_APPROVAL = "Pending for Validator Approval"
    STATUS_PENDING_HOD_APPROVAL = "Pending for HOD Approval"
    STATUS_PENDING_SAP_CODE_CREATION = "Pending for SAP Code Creation"

    def extract_pan_from_gstin(value: Any) -> str:
        raw = str(value or "").strip().upper()
        if not raw:
            return ""
        cleaned = "".join(ch for ch in raw if ch.isalnum())
        if len(cleaned) < 12:
            return ""
        pan_candidate = cleaned[2:12]
        if m.PAN_REGEX.fullmatch(pan_candidate):
            return pan_candidate
        return ""

    def customer_now_ist_iso() -> str:
        ist = timezone(timedelta(hours=5, minutes=30))
        return datetime.now(ist).isoformat(timespec="seconds")

    def normalize_inline_document_content_type(file_name: str, content_type: Any) -> str:
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


    def sniff_document_content_type(file_bytes: bytes) -> str:
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


    def resolve_inline_document_content_type(file_name: str, content_type: Any, file_bytes: bytes) -> str:
        sniffed_content_type = sniff_document_content_type(file_bytes)
        if sniffed_content_type:
            return sniffed_content_type
        return normalize_inline_document_content_type(file_name, content_type)


    def build_customer_inline_document_response(
        file_name: str,
        content: Any,
        content_type: Any,
    ) -> Response:
        normalized_name = m._extract_document_filename(file_name) or Path(str(file_name or "document")).name or "document"
        file_bytes = bytes(content or b"")
        resolved_content_type = resolve_inline_document_content_type(normalized_name, content_type, file_bytes)
        headers = {"Content-Disposition": f'inline; filename="{normalized_name}"'}
        return Response(content=file_bytes, media_type=resolved_content_type, headers=headers)

    def decode_reference_text_payload(file_bytes: bytes) -> str:
        payload = bytes(file_bytes or b"")
        if not payload:
            return ""
        if len(payload) > 4096:
            return ""

        def reference_candidates(raw_text: str) -> list[str]:
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

            for candidate in reference_candidates(text):
                lowered = candidate.lower()
                if lowered.startswith(("http://", "https://", "/mock-files/")):
                    return candidate
                if "/" in candidate or "\\" in candidate:
                    return candidate
                if len(candidate) <= 256 and all(ch.isalnum() or ch in "-_. " for ch in candidate):
                    return candidate

        return ""

    def looks_like_reference_text_payload(file_bytes: bytes) -> bool:
        return bool(decode_reference_text_payload(file_bytes))

    def has_valid_document_signature(file_name: str, file_bytes: bytes) -> bool:
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
        if looks_like_reference_text_payload(payload):
            return False
        return True

    def resolve_customer_document_reference(reference_value: Any) -> Response | None:
        raw_reference = str(reference_value or "").strip()
        if not raw_reference:
            return None

        normalized_reference = raw_reference
        if normalized_reference.startswith("/mock-files/"):
            normalized_reference = normalized_reference.split("/mock-files/", 1)[1]

        storage_key, normalized_name = m._normalize_mock_file_reference(normalized_reference)
        if not normalized_name:
            return None

        stored = m.mock_files.get(storage_key) or m.mock_files.get(normalized_name)
        if stored:
            return build_customer_inline_document_response(
                normalized_name,
                stored.get("content") or b"",
                stored.get("content_type"),
            )

        disk_path = m._resolve_uploaded_file_disk_path(storage_key)
        if disk_path is not None and disk_path.exists() and disk_path.is_file():
            guessed_type = mimetypes.guess_type(str(disk_path))[0]
            return build_customer_inline_document_response(
                disk_path.name,
                disk_path.read_bytes(),
                guessed_type,
            )

        disk_path = m._find_uploaded_file_by_name(normalized_name)
        if disk_path is not None:
            guessed_type = mimetypes.guess_type(str(disk_path))[0]
            return build_customer_inline_document_response(
                disk_path.name,
                disk_path.read_bytes(),
                guessed_type,
            )

        return None

    def normalize_status_key(value: Any) -> str:
        return " ".join(str(value or "").strip().lower().split())

    def is_pending_validator_status(value: Any) -> bool:
        normalized = normalize_status_key(value)
        if not normalized:
            return False
        if normalized in {"pending for validator approval", "pendign for validator approval"}:
            return True
        return "pending" in normalized and ("validator" in normalized or "validation" in normalized)

    def is_pending_hod_status(value: Any) -> bool:
        normalized = normalize_status_key(value)
        if not normalized:
            return False
        if normalized in {"pending for hod approval", "pendign for hod approval"}:
            return True
        return "pending" in normalized and "hod" in normalized

    def is_pending_user_status(value: Any) -> bool:
        normalized = normalize_status_key(value)
        if not normalized:
            return False
        if normalized in {
            "pending for user approval",
            "pendign for user approval",
            "pending for buyer approval",
            "pendign for buyer approval",
            "pending for data validation",
        }:
            return True
        return "pending" in normalized and ("user" in normalized or "buyer" in normalized)

    def normalize_customer_reject_target(value: Any) -> str:
        raw = str(m.empty_to_none(value) or "").strip().lower()
        if raw in {"customer", "vendor"}:
            return "customer"
        if raw == "buyer":
            return "buyer"
        return ""

    # Kept for backward compatibility; no GST fields are obsolete now.
    OBSOLETE_CUSTOMER_GST_FIELDS: set[str] = set()
    OBSOLETE_CUSTOMER_GST_DOC_FIELDS: set[str] = set()
    REMOVED_UNUSED_CUSTOMER_FIELDS = {
        "validateby",
        "validate by",
        "msmevalidate",
        "msme validate",
        "pfvalidate",
        "pf validate",
        "esivalidate",
        "esi validate",
        "bankvalidate",
        "bank validate",
        "companyname",
        "registrationpath",
        "registrationlink",
        "sqlrecordid",
        "reviewdecision",
        "rejectto",
        "reject to",
        "invitedbyid",
        "invitedbyemail",
        "invited by email",
        "buyeremail",
        "contactpersonemail",
        "contactemailid",
        "rubaminapproverhod",
        "rubaminapprovalhod",
    }

    def is_obsolete_customer_field(field_name: str) -> bool:
        key = str(field_name or "").strip().lower()
        if not key:
            return False
        if key in OBSOLETE_CUSTOMER_GST_FIELDS:
            return True
        if key in REMOVED_UNUSED_CUSTOMER_FIELDS:
            return True

        for suffix in ("_filename", "_contenttype", "_filesize", "filename", "contenttype", "filesize"):
            if not key.endswith(suffix):
                continue
            base_key = key[: -len(suffix)].rstrip("_")
            if base_key in OBSOLETE_CUSTOMER_GST_DOC_FIELDS:
                return True
        return False

    def sanitize_customer_updates(payload: dict[str, Any] | None) -> dict[str, Any]:
        if not isinstance(payload, dict):
            return {}
        return {
            str(key): value
            for key, value in payload.items()
            if str(key or "").strip() and not is_obsolete_customer_field(str(key))
        }

    def normalize_public_base_url(raw_base: Any) -> str:
        base = str(raw_base or "").strip().rstrip("/")
        if not base or "://" not in base:
            return ""
        parsed = urlparse(base)
        if not parsed.scheme or not parsed.netloc:
            return ""
        normalized = f"{parsed.scheme}://{parsed.netloc}{parsed.path}".rstrip("/")
        return normalized

    def build_customer_public_url(path: str, request: Request | None = None) -> str:
        normalized_path = "/" + str(path or "").strip().lstrip("/")
        for candidate in (
            getattr(m, "PUBLIC_BASE_URL", ""),
            str(request.base_url or "").strip().rstrip("/") if request is not None else "",
        ):
            normalized_base = normalize_public_base_url(candidate)
            if normalized_base:
                return f"{normalized_base}{normalized_path}"
        return normalized_path

    def build_customer_dashboard_url(request: Request | None = None) -> str:
        return build_customer_public_url("/web/customer-dashboard", request=request)

    @app.get("/web/customer-public-form", response_class=HTMLResponse)
    def web_customer_public_form() -> HTMLResponse:
        return HTMLResponse(content=m.read_template("customer_public_form.html"))

    @app.get("/api/customer/get-hod-emails")
    def get_customer_hod_emails(
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> list[dict[str, Any]]:
        normalized_role = m.normalize_role(user.get("role"))
        bot_name = m.normalize_bot_name(user.get("botName") or "Vendor Bot")
        if normalized_role != "admin" and bot_name == "Vendor Bot":
            raise HTTPException(status_code=403, detail="Access denied for Vendor Bot")

        allowed_bot_names = {"Both", "Customer Bot"}
        merged: dict[str, dict[str, Any]] = {}

        def add_hod_candidate(candidate: dict[str, Any], hydrate_local: bool = False) -> None:
            if not isinstance(candidate, dict):
                return
            role_value = m.normalize_role(candidate.get("role"))
            if role_value != "hod":
                return

            normalized_bot = m.normalize_bot_name(candidate.get("botName") or "Vendor Bot")
            if normalized_bot not in allowed_bot_names:
                return

            email = str(candidate.get("email") or "").strip()
            if not email:
                return

            name = str(candidate.get("name") or email).strip() or email
            user_id = candidate.get("id")
            if hydrate_local:
                local_user = m.ensure_local_user(
                    name=name,
                    email=email,
                    role="HOD",
                    preferred_id=user_id,
                    bot_name=normalized_bot,
                )
                user_id = local_user.get("id")
                name = str(local_user.get("name") or name).strip() or email

            merged[email.lower()] = {
                "id": user_id,
                "name": name,
                "email": email,
            }

        for local_user in m.users.values():
            add_hod_candidate(local_user)

        for db_user in m.get_db_users():
            add_hod_candidate(db_user, hydrate_local=True)

        values = list(merged.values())
        values.sort(key=lambda item: (str(item.get("name") or "").lower(), str(item.get("email") or "").lower()))
        return values

    @app.post("/api/customer/db/remove-obsolete-gst-columns")
    def remove_customer_obsolete_gst_columns(
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        if m.normalize_role(user.get("role")) != "admin":
            raise HTTPException(status_code=403, detail="Only Admin can remove obsolete customer columns")
        result = customer_db.remove_obsolete_customer_gst_columns()
        return {
            "message": "Obsolete customer GST 2/GST 3 columns cleanup completed",
            "removedColumns": result.get("removedColumns", []),
            "missingColumns": result.get("missingColumns", []),
            "table": result.get("table"),
        }

    @app.post("/api/customer/db/cleanup-duplicate-columns")
    def cleanup_customer_duplicate_columns(
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        if m.normalize_role(user.get("role")) != "admin":
            raise HTTPException(status_code=403, detail="Only Admin can cleanup duplicate customer columns")

        result = customer_db.cleanup_customer_duplicate_columns()
        return {
            "message": (
                "Customer duplicate column cleanup completed"
                if result.get("success")
                else "Customer duplicate column cleanup failed"
            ),
            "success": bool(result.get("success")),
            "schema": result.get("schema"),
            "table": result.get("table"),
            "droppedColumns": result.get("droppedColumns", []),
            "droppedDuplicateColumns": result.get("droppedDuplicateColumns", []),
            "droppedRemovedColumns": result.get("droppedRemovedColumns", []),
            "error": result.get("error"),
        }

    @app.get("/api/customer/dashboard-records")
    def get_customer_dashboard_records(
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> list[dict[str, Any]]:
        normalized_role = m.normalize_role(user.get("role"))
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
                for row in get_customer_dashboard_records(scoped_user):
                    row_key = str(row.get("record") or "").strip().lower()
                    if not row_key:
                        row_key = f"temp-{len(merged_rows) + 1}"
                    existing_row = merged_rows.get(row_key)
                    if not existing_row or str(row.get("createdAt") or "") >= str(existing_row.get("createdAt") or ""):
                        merged_rows[row_key] = row
            rows = list(merged_rows.values())
            rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
            return rows

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

        sql_rows = load_customer_sql_records()

        if normalized_role == "admin":
            for invite in iter_customer_invite_records():
                upsert_row(map_customer_summary(invite))

            for form in iter_customer_local_form_records():
                upsert_row(map_customer_summary(form))

            for row in sql_rows:
                upsert_row(map_customer_summary(row))

            return build_rows()

        if normalized_role == "hod":
            hod_email = str(user.get("email") or "").strip().lower()
            if not hod_email:
                return []

            def is_assigned_to_hod(record: dict[str, Any] | None) -> bool:
                return m._extract_record_hod_email(record) == hod_email

            for invite in iter_customer_invite_records():
                status_value = get_customer_status_value(invite)
                if not is_pending_hod_status(status_value):
                    continue
                if not is_assigned_to_hod(invite):
                    continue
                upsert_row(map_customer_summary(invite))

            for form in iter_customer_local_form_records():
                if not is_pending_hod_status(get_customer_status_value(form)):
                    continue
                if not is_assigned_to_hod(form):
                    continue
                upsert_row(map_customer_summary(form))

            for row in sql_rows:
                status_value = get_customer_status_value(row)
                if not is_pending_hod_status(status_value):
                    continue
                if not is_assigned_to_hod(row):
                    continue
                upsert_row(map_customer_summary(row))

            return build_rows()

        if normalized_role == "validator":
            for invite in iter_customer_invite_records():
                if not is_pending_validator_status(get_customer_status_value(invite)):
                    continue
                upsert_row(map_customer_summary(invite))

            for form in iter_customer_local_form_records():
                if not is_pending_validator_status(get_customer_status_value(form)):
                    continue
                upsert_row(map_customer_summary(form))

            for row in sql_rows:
                status_value = get_customer_status_value(row)
                if not is_pending_validator_status(status_value):
                    continue
                upsert_row(map_customer_summary(row))

            return build_rows()

        user_email = str(user.get("email") or "").strip().lower()
        if not user_email:
            return []

        for invite in iter_customer_invite_records():
            if normalized_role in {"user", "buyer"}:
                is_assigned = is_customer_user_assigned_to_rubamin_contact(invite, user_email)
            else:
                is_assigned = m._record_is_assigned_to_email(invite, user_email)
            if not is_assigned:
                continue
            upsert_row(map_customer_summary(invite))

        for form in iter_customer_local_form_records():
            if normalized_role in {"user", "buyer"}:
                is_assigned = is_customer_user_assigned_to_rubamin_contact(form, user_email)
            else:
                is_assigned = m._record_is_assigned_to_email(form, user_email)
            if not is_assigned:
                continue
            upsert_row(map_customer_summary(form))

        for row in sql_rows:
            if normalized_role in {"user", "buyer"}:
                is_assigned = is_customer_user_assigned_to_rubamin_contact(row, user_email)
            else:
                is_assigned = m._record_is_assigned_to_email(row, user_email)
            if not is_assigned:
                continue
            upsert_row(map_customer_summary(row))

        assigned_rows = build_rows()
        if assigned_rows or normalized_role not in {"user", "buyer"}:
            return assigned_rows

        # Fallback: if assignment metadata is missing on older records, still show
        # actionable user-stage rows instead of an empty table.
        for invite in iter_customer_invite_records():
            if not is_pending_user_status(get_customer_status_value(invite)):
                continue
            upsert_row(map_customer_summary(invite))

        for form in iter_customer_local_form_records():
            if not is_pending_user_status(get_customer_status_value(form)):
                continue
            upsert_row(map_customer_summary(form))

        for row in sql_rows:
            if not is_pending_user_status(get_customer_status_value(row)):
                continue
            upsert_row(map_customer_summary(row))

        return build_rows()

    @app.get("/api/customer/history-records")
    def get_customer_history_records(
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> list[dict[str, Any]]:
        normalized_role = m.normalize_role(user.get("role"))
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
                for row in get_customer_history_records(scoped_user):
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
        if not user_email and normalized_role != "admin":
            return []

        merged: dict[str, dict[str, Any]] = {}

        def upsert(row: dict[str, Any]) -> None:
            key = str(row.get("record") or "").strip().lower()
            if not key:
                return
            existing = merged.get(key)
            if not existing or str(row.get("createdAt") or "") >= str(existing.get("createdAt") or ""):
                merged[key] = row

        def walk_sources(collector) -> None:
            for invite in iter_customer_invite_records():
                collector(invite)
            for form in iter_customer_local_form_records():
                collector(form)
            for row in load_customer_sql_records():
                collector(row)

        if normalized_role == "admin":
            walk_sources(lambda record: upsert(map_customer_summary(record)))
        elif normalized_role == "validator":
            # Match vendor behavior where validator can review history without strict assignee dependency.
            walk_sources(lambda record: upsert(map_customer_summary(record)))
        elif normalized_role == "hod":
            def collect_hod(record: dict[str, Any]) -> None:
                if m._extract_record_hod_email(record) != user_email:
                    return
                upsert(map_customer_summary(record))

            walk_sources(collect_hod)
        elif normalized_role in {"user", "buyer"}:
            def collect_user(record: dict[str, Any]) -> None:
                if is_customer_user_assigned_to_rubamin_contact(record, user_email):
                    upsert(map_customer_summary(record))

            walk_sources(collect_user)

            if not merged:
                # Fallback for legacy records missing assignee fields.
                def collect_pending_user(record: dict[str, Any]) -> None:
                    if is_pending_user_status(get_customer_status_value(record)):
                        upsert(map_customer_summary(record))

                walk_sources(collect_pending_user)
        else:
            def collect_default(record: dict[str, Any]) -> None:
                if m._record_is_assigned_to_email(record, user_email):
                    upsert(map_customer_summary(record))

            walk_sources(collect_default)

        rows = list(merged.values())
        rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
        return rows

    @app.get("/api/customer/shipto-records")
    def get_customer_shipto_records(
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> list[dict[str, Any]]:
        if not m.user_has_any_role(user, "user", "buyer"):
            raise HTTPException(status_code=403, detail="Only User role can access Ship-to records")

        user_email = str(user.get("email") or "").strip().lower()
        if not user_email:
            return []

        scoped_user = dict(user)
        scoped_user["role"] = "user"
        rows = get_customer_history_records(scoped_user)

        merged: dict[str, dict[str, Any]] = {}
        for row in rows:
            record_id = str(row.get("record") or "").strip()
            if not record_id:
                continue

            local_record, sql_record = get_customer_records(record_id)
            assigned_to_login_user = (
                is_customer_user_assigned_to_rubamin_contact(local_record, user_email)
                or is_customer_user_assigned_to_rubamin_contact(sql_record, user_email)
                or m._record_is_assigned_to_email(local_record, user_email)
                or m._record_is_assigned_to_email(sql_record, user_email)
            )
            if not assigned_to_login_user:
                continue

            row_key = record_id.lower()
            summary = {
                "record": record_id,
                "company": str(row.get("company") or "").strip(),
                "name": str(row.get("name") or "").strip(),
                "category": str(row.get("category") or "").strip(),
                "approverStatus": str(row.get("approverStatus") or "").strip(),
                "type": str(row.get("type") or "").strip(),
                "createdAt": str(row.get("createdAt") or "").strip(),
            }
            existing = merged.get(row_key)
            if not existing or summary["createdAt"] >= str(existing.get("createdAt") or ""):
                merged[row_key] = summary

        result = list(merged.values())
        result.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
        return result

    def get_customer_history_record_export_context(
        user: dict[str, Any],
        record_id: str,
    ) -> tuple[str, dict[str, Any], dict[str, Any] | None, dict[str, Any] | None]:
        normalized_role = m.normalize_role(user.get("role"))
        user_email = str(user.get("email") or "").strip().lower()
        if not m.user_has_any_role(user, "admin") and not user_email:
            raise HTTPException(status_code=403, detail="Unable to identify logged-in user")

        requested_record_id = m._normalize_record_id(record_id)
        if not requested_record_id:
            raise HTTPException(status_code=400, detail="recordId is required")

        summary_rows = get_customer_history_records(user)
        summary_row_by_key: dict[str, dict[str, Any]] = {}
        for row in summary_rows:
            row_record_id = m._normalize_record_id(row.get("record"))
            if not row_record_id:
                continue
            row_key = row_record_id.lower()
            if row_key not in summary_row_by_key:
                summary_row_by_key[row_key] = row

        requested_key = requested_record_id.lower()
        summary_row = summary_row_by_key.get(requested_key)

        local_record, sql_record = get_customer_records(requested_record_id)
        if not local_record and not sql_record and not summary_row:
            raise HTTPException(status_code=404, detail="Customer record not found")

        if m.user_has_any_role(user, "admin", "validator"):
            is_allowed = True
        elif m.user_has_any_role(user, "hod"):
            allowed_hod_emails = {
                email
                for email in (
                    m._extract_record_hod_email(local_record),
                    m._extract_record_hod_email(sql_record),
                )
                if email
            }
            is_allowed = bool(user_email and user_email in allowed_hod_emails)
        elif m.user_has_any_role(user, "user", "buyer"):
            is_allowed = bool(
                summary_row
                or is_customer_user_assigned_to_rubamin_contact(local_record, user_email)
                or is_customer_user_assigned_to_rubamin_contact(sql_record, user_email)
            )
        else:
            is_allowed = bool(
                summary_row
                or m._record_is_assigned_to_email(local_record, user_email)
                or m._record_is_assigned_to_email(sql_record, user_email)
            )

        if not is_allowed:
            raise HTTPException(status_code=403, detail="Access denied for this record")

        merged_record = merge_customer_records(requested_record_id, local_record, sql_record)
        if not merged_record and summary_row:
            merged_record = {
                "record": str(summary_row.get("record") or requested_record_id),
                "company": str(summary_row.get("company") or ""),
                "category": str(summary_row.get("category") or ""),
                "approverStatus": str(summary_row.get("approverStatus") or ""),
                "name": str(summary_row.get("name") or ""),
                "type": str(summary_row.get("type") or ""),
                "createdAt": str(summary_row.get("createdAt") or ""),
                "recordId": requested_record_id,
            }

        merged_record["recordId"] = requested_record_id
        return requested_record_id, merged_record, local_record, sql_record

    def build_customer_export_filename_base(record_id: str, record: dict[str, Any] | None) -> str:
        safe_record = m._sanitize_download_filename_part(record_id, "Record")
        display_name = (
            get_customer_display_name(record)
            or pick_first(record, "companyName", "CompanyName", "company", "Company", "name", "Name")
            or "Customer"
        )
        safe_name = m._sanitize_download_filename_part(display_name, "Customer")
        return f"{safe_record} + {safe_name}"

    def build_customer_single_record_excel_bytes(record: dict[str, Any]) -> bytes:
        filtered_record = m._filter_record_for_history_export(record or {})

        workbook = m.Workbook()
        worksheet = workbook.active
        worksheet.title = "Customer Record"

        if not filtered_record:
            worksheet.append(["No records found"])
            worksheet.column_dimensions["A"].width = 28
        else:
            headers = list(filtered_record.keys())
            worksheet.append(headers)
            values: list[Any] = []
            column_widths = [len(str(header or "")) for header in headers]

            for idx, header in enumerate(headers):
                cell_value = m._coerce_excel_cell_value(filtered_record.get(header))
                values.append(cell_value)
                text_length = len(str(cell_value or ""))
                if text_length > column_widths[idx]:
                    column_widths[idx] = text_length

            worksheet.append(values)
            for idx, width in enumerate(column_widths, start=1):
                worksheet.column_dimensions[m.get_column_letter(idx)].width = min(80, max(12, width + 2))
            worksheet.freeze_panes = "A2"

        output = m.BytesIO()
        workbook.save(output)
        workbook.close()
        return output.getvalue()

    def build_customer_history_pdf_bytes(record_id: str, record: dict[str, Any]) -> bytes:
        filtered_record = m._filter_record_for_history_export(record or {})
        if not filtered_record:
            filtered_record = {
                "Record ID": str(record_id or ""),
                "Message": "No data available",
            }

        rows_html = "".join(
            (
                f"<tr><th>{m.html_escape(str(key))}</th>"
                f"<td>{m.html_escape(str(value if value is not None else ''))}</td></tr>"
            )
            for key, value in filtered_record.items()
        )
        html_content = (
            "<!doctype html>"
            "<html><head><meta charset='utf-8'/>"
            "<style>"
            "body{font-family:Arial,sans-serif;padding:24px;color:#0f172a;}"
            "h1{font-size:22px;margin:0 0 14px;color:#1e3a8a;}"
            "table{width:100%;border-collapse:collapse;table-layout:fixed;}"
            "th,td{border:1px solid #cbd5e1;padding:8px 10px;vertical-align:top;word-break:break-word;}"
            "th{background:#e2e8f0;text-align:left;width:34%;}"
            "</style></head><body>"
            f"<h1>Customer Record - {m.html_escape(str(record_id or '-'))}</h1>"
            f"<table>{rows_html}</table>"
            "</body></html>"
        )
        return m._render_html_to_pdf_with_browser(html_content)

    def get_customer_document_field_candidates(*records: dict[str, Any] | None) -> list[str]:
        candidates: list[str] = []
        seen: set[str] = set()

        def add_candidate(value: Any) -> None:
            text = str(value or "").strip()
            if not text:
                return
            marker = text.lower()
            if marker in seen:
                return
            seen.add(marker)
            candidates.append(text)

        for field_name in m.DOCUMENT_FIELDS:
            add_candidate(field_name)
        for field_name in m.IMPORT_FORM_UPLOAD_TO_COLUMN.keys():
            add_candidate(field_name)

        document_hints = ("file", "filename", "document", "attachment", "upload", "certificate", "cert", "doc")
        for record in records:
            if not isinstance(record, dict):
                continue
            for key in record.keys():
                key_text = str(key or "").strip()
                if not key_text:
                    continue
                lowered_key = key_text.lower()
                if not any(hint in lowered_key for hint in document_hints):
                    continue
                add_candidate(key_text)

                base_name = key_text
                if lowered_key.endswith("_filename"):
                    base_name = key_text[: -len("_FileName")]
                elif lowered_key.endswith("filename"):
                    base_name = key_text[: -len("FileName")]
                base_name = base_name.rstrip("_")
                if base_name:
                    add_candidate(base_name)

        return candidates

    @app.get("/api/customer/history-records/export")
    def export_customer_history_record(
        recordId: str,
        exportType: str,
        user: dict[str, Any] = Depends(m.get_current_user),
    ):
        normalized_export_type = str(exportType or "").strip().lower()
        if normalized_export_type not in {"excel", "pdf", "documents"}:
            raise HTTPException(status_code=400, detail="exportType must be excel, pdf, or documents")

        record_id, merged_record, local_record, sql_record = get_customer_history_record_export_context(user, recordId)
        filename_base = build_customer_export_filename_base(record_id, merged_record)

        if normalized_export_type == "excel":
            excel_bytes = build_customer_single_record_excel_bytes(merged_record)
            headers_map = {"Content-Disposition": f'attachment; filename="{filename_base}.xlsx"'}
            return Response(
                content=excel_bytes,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers=headers_map,
            )

        if normalized_export_type == "pdf":
            try:
                pdf_bytes = build_customer_history_pdf_bytes(record_id, merged_record)
            except RuntimeError as exc:
                raise HTTPException(status_code=500, detail=str(exc))
            headers_map = {"Content-Disposition": f'attachment; filename="{filename_base}.pdf"'}
            return Response(content=pdf_bytes, media_type="application/pdf", headers=headers_map)

        document_fields = get_customer_document_field_candidates(merged_record, local_record, sql_record)
        collected_documents: list[dict[str, Any]] = []
        seen_content_hashes: set[str] = set()
        for index, field_name in enumerate(document_fields, start=1):
            try:
                document_response = get_customer_document(recordId=record_id, fieldName=field_name, user=user)
            except HTTPException as exc:
                if exc.status_code in {400, 403, 404}:
                    continue
                raise

            content_bytes, _, response_file_name = m._extract_document_response_payload(document_response)
            if not content_bytes:
                continue

            content_hash = m.hashlib.sha1(content_bytes).hexdigest()
            if content_hash in seen_content_hashes:
                continue
            seen_content_hashes.add(content_hash)

            normalized_file_name = m._extract_document_filename(response_file_name) or f"{field_name}.bin"
            collected_documents.append(
                {
                    "fallback_name": f"document_{index}",
                    "filename": normalized_file_name,
                    "content": content_bytes,
                }
            )

        if not collected_documents:
            raise HTTPException(status_code=404, detail="No uploaded documents found for this record")

        output = m.BytesIO()
        used_entry_names: set[str] = set()
        with m.zipfile.ZipFile(output, mode="w", compression=m.zipfile.ZIP_DEFLATED) as zip_file:
            for item in collected_documents:
                entry_name = m._build_unique_zip_entry_name(
                    file_name=item.get("filename") or "",
                    used_names=used_entry_names,
                    fallback_prefix=str(item.get("fallback_name") or "document"),
                )
                zip_file.writestr(entry_name, item.get("content") or b"")
        output.seek(0)

        headers_map = {"Content-Disposition": f'attachment; filename="{filename_base}.zip"'}
        return m.StreamingResponse(output, media_type="application/zip", headers=headers_map)

    @app.get("/api/customer/history-records/download")
    def download_customer_history_records_excel(
        user: dict[str, Any] = Depends(m.get_current_user),
        recordId: str | None = None,
    ) -> Response:
        normalized_role = m.normalize_role(user.get("role"))
        user_email = str(user.get("email") or "").strip().lower()
        if normalized_role != "admin" and not user_email:
            raise HTTPException(status_code=403, detail="Unable to identify logged-in user")

        requested_record_id = m._normalize_record_id(recordId)
        if recordId is not None and not requested_record_id:
            raise HTTPException(status_code=400, detail="recordId is invalid")

        summary_rows = get_customer_history_records(user)
        summary_rows.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)

        record_ids: list[str] = []
        seen_record_ids: set[str] = set()
        record_id_by_key: dict[str, str] = {}
        summary_row_by_record_id: dict[str, dict[str, Any]] = {}
        for db_row in summary_rows:
            row_record_id = m._normalize_record_id(db_row.get("record"))
            if not row_record_id:
                continue
            row_key = row_record_id.lower()
            if row_key in seen_record_ids:
                continue
            seen_record_ids.add(row_key)
            record_ids.append(row_record_id)
            record_id_by_key[row_key] = row_record_id
            summary_row_by_record_id[row_key] = db_row

        if requested_record_id:
            requested_key = requested_record_id.lower()
            if requested_key in record_id_by_key:
                record_ids = [record_id_by_key[requested_key]]
            else:
                get_customer_history_record_export_context(user, requested_record_id)
                record_ids = [requested_record_id]

        export_rows: list[dict[str, Any]] = []
        headers: list[str] = []
        seen_headers: set[str] = set()

        for row_record_id in record_ids:
            row_key = row_record_id.lower()
            local_record, sql_record = get_customer_records(row_record_id)
            full_record = merge_customer_records(row_record_id, local_record, sql_record)

            if not full_record:
                fallback_row = summary_row_by_record_id.get(row_key) or {}
                full_record = {
                    "record": str(fallback_row.get("record") or row_record_id),
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
                if not column_name or m._is_document_column_for_history_export(column_name):
                    continue
                filtered_record[column_name] = value
                if column_name not in seen_headers:
                    seen_headers.add(column_name)
                    headers.append(column_name)

            if filtered_record:
                export_rows.append(filtered_record)

        workbook = m.Workbook()
        worksheet = workbook.active
        worksheet.title = "Customer History"

        if not headers:
            worksheet.append(["No records found"])
            worksheet.column_dimensions["A"].width = 28
        else:
            worksheet.append(headers)
            column_widths = [len(str(header or "")) for header in headers]

            for row in export_rows:
                normalized_row = {
                    m._normalize_history_export_column_key(key): value
                    for key, value in row.items()
                    if str(key or "").strip()
                }
                row_values: list[Any] = []
                for idx, header in enumerate(headers):
                    if header in row:
                        raw_value = row.get(header)
                    else:
                        normalized_header = m._normalize_history_export_column_key(header)
                        raw_value = normalized_row.get(normalized_header)
                    cell_value = m._coerce_excel_cell_value(raw_value)
                    row_values.append(cell_value)

                    text_length = len(str(cell_value or ""))
                    if text_length > column_widths[idx]:
                        column_widths[idx] = text_length

                worksheet.append(row_values)

            for idx, width in enumerate(column_widths, start=1):
                worksheet.column_dimensions[m.get_column_letter(idx)].width = min(80, max(12, width + 2))
            worksheet.freeze_panes = "A2"

        output = m.BytesIO()
        workbook.save(output)
        workbook.close()
        excel_bytes = output.getvalue()

        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d_%H-%M-%S")
        if requested_record_id:
            filename = f"Customer Record ({timestamp}).xlsx"
        else:
            filename = f"Customer List ({timestamp}).xlsx"

        headers_map = {"Content-Disposition": f'attachment; filename="{filename}"'}
        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers_map,
        )

    @app.get("/api/customer/record")
    def get_customer_record_by_id(
        recordId: str,
        user: dict[str, Any] | None = Depends(m.get_optional_current_user),
    ) -> dict[str, Any]:
        record_id = m._normalize_record_id(recordId)
        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Customer record not found")

        authorize_customer_record_access(record_id, local_record, sql_record, user)
        merged_record = merge_customer_records(record_id, local_record, sql_record)

        pre_screen_answers: dict[str, str] = {}
        pre_screen_answers.update(m._extract_prescreen_answers(sql_record))
        pre_screen_answers.update(m._extract_prescreen_answers(local_record))

        return {
            "recordId": record_id,
            "vendor": merged_record,  # kept for frontend compatibility
            "preScreenAnswers": pre_screen_answers,
        }

    @app.get("/api/customer/gst-filing-table")
    def get_customer_gst_filing_table(
        recordId: str,
        financialYear: str | None = None,
        user: dict[str, Any] | None = Depends(m.get_optional_current_user),
    ) -> dict[str, Any]:
        record_id = m._normalize_record_id(recordId)
        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Customer record not found")

        authorize_customer_record_access(record_id, local_record, sql_record, user)

        approver_status = m._first_existing_record_value(
            local_record,
            sql_record,
            "approverStatus",
            "ApproverStatus",
            "ApprovStatus",
            "status",
            "Status",
        )
        if not is_pending_validator_status(approver_status):
            raise HTTPException(status_code=400, detail="GST filing table is available only for Pending For Validator Approval records")

        gst_status = m._first_existing_record_value(
            local_record,
            sql_record,
            "GSTStatus",
            "gstStatus",
            "gstRegistrationStatus",
        )
        if not m._is_registered_gst_status(gst_status):
            raise HTTPException(status_code=400, detail="GST filing table is available only for GST registered records")

        gst_number = str(
            m._first_existing_record_value(
                local_record,
                sql_record,
                "GSTIN",
                "gstin",
                "gstNumber",
            )
            or ""
        ).strip().upper()
        if not gst_number:
            raise HTTPException(status_code=400, detail="GST number not found for this record")

        normalized_financial_year = m._normalize_financial_year_value(financialYear)
        result = m._run_gst_filing_table_extraction_inline(
            gst_number,
            financial_year=normalized_financial_year or None,
        )
        if not isinstance(result, dict):
            raise HTTPException(status_code=500, detail="Unable to extract GST filing tables")

        return {
            "recordId": record_id,
            "gstNumber": gst_number,
            "financialYear": normalized_financial_year,
            "result": result,
        }

    @app.get("/api/customer/document")
    def get_customer_document(
        recordId: str,
        fieldName: str,
        user: dict[str, Any] | None = Depends(m.get_optional_current_user),
    ) -> Response:
        record_id = m._normalize_record_id(recordId)
        field_name = str(fieldName or "").strip()
        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")
        if not field_name:
            raise HTTPException(status_code=400, detail="fieldName is required")

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Customer record not found")
        authorize_customer_record_access(record_id, local_record, sql_record, user)

        base_candidate_keys = m._get_document_candidate_keys(field_name)
        canonical_field_name = base_candidate_keys[0] if base_candidate_keys else field_name

        candidate_keys: list[str] = []
        seen_candidate_keys: set[str] = set()

        def add_candidate_key(key_name: str) -> None:
            normalized_key = str(key_name or "").strip()
            if not normalized_key:
                return
            marker = normalized_key.lower()
            if marker in seen_candidate_keys:
                return
            seen_candidate_keys.add(marker)
            candidate_keys.append(normalized_key)

        for key_name in base_candidate_keys:
            base_key = str(key_name or "").strip()
            if not base_key:
                continue
            add_candidate_key(base_key)

            lowered = base_key.lower()
            if lowered.endswith("_filename"):
                base_key = base_key[: -len("_FileName")]
            elif lowered.endswith("filename"):
                base_key = base_key[: -len("FileName")]
            base_key = base_key.rstrip("_")
            if not base_key:
                continue
            add_candidate_key(base_key)
            add_candidate_key(f"{base_key}_FileName")
            add_candidate_key(f"{base_key}FileName")

        if not candidate_keys:
            candidate_keys = [canonical_field_name]

        raw_reference_candidates: list[str] = []
        seen_references: set[str] = set()
        for record in (sql_record, local_record):
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
            extracted_name = m._extract_document_filename(raw_reference)
            if not extracted_name:
                continue
            dedupe_key = extracted_name.lower()
            if dedupe_key in seen_file_names:
                continue
            seen_file_names.add(dedupe_key)
            extracted_file_names.append(extracted_name)

        disk_path: Path | None = None
        field_name_candidates = [canonical_field_name, *base_candidate_keys[1:]]
        for field_key in field_name_candidates:
            for extracted_name in extracted_file_names:
                disk_path = m._find_uploaded_file_for_record_field(record_id, field_key, extracted_name)
                if disk_path is not None:
                    break
            if disk_path is not None:
                break
            disk_path = m._find_uploaded_file_for_record_field(record_id, field_key)
            if disk_path is not None:
                break
        if disk_path is not None:
            guessed_type = mimetypes.guess_type(str(disk_path))[0]
            return build_customer_inline_document_response(
                disk_path.name,
                disk_path.read_bytes(),
                guessed_type,
            )

        stored_document = customer_db.get_customer_document_content_by_record_id(record_id, candidate_keys)
        if stored_document:
            file_name = m._extract_document_filename(stored_document.get("filename"))
            preferred_extracted_name = next(
                (name for name in extracted_file_names if Path(str(name)).suffix),
                extracted_file_names[0] if extracted_file_names else "",
            )
            if (
                not file_name
                or not Path(file_name).suffix
                or file_name.lower().endswith(".bin")
            ) and preferred_extracted_name:
                file_name = preferred_extracted_name
            if not file_name:
                file_name = "document"

            stored_content = bytes(stored_document.get("content") or b"")
            if looks_like_reference_text_payload(stored_content):
                decoded_reference = decode_reference_text_payload(stored_content)
                if decoded_reference:
                    marker = decoded_reference.lower()
                    if marker not in seen_references:
                        seen_references.add(marker)
                        raw_reference_candidates.append(decoded_reference)
                        decoded_name = m._extract_document_filename(decoded_reference)
                        if decoded_name:
                            decoded_name_key = decoded_name.lower()
                            if decoded_name_key not in seen_file_names:
                                seen_file_names.add(decoded_name_key)
                                extracted_file_names.append(decoded_name)
            elif has_valid_document_signature(file_name, stored_content):
                return build_customer_inline_document_response(
                    file_name,
                    stored_content,
                    stored_document.get("content_type"),
                )

        for extracted_name in extracted_file_names:
            resolved_response = resolve_customer_document_reference(extracted_name)
            if resolved_response is not None:
                return resolved_response

        # Fallback for legacy references.
        for raw_reference in raw_reference_candidates:
            mock_reference = raw_reference
            lower_reference = mock_reference.lower()
            if lower_reference.startswith("http://") or lower_reference.startswith("https://"):
                proxied_remote_response = m._build_inline_document_response_from_remote_url(mock_reference)
                if proxied_remote_response is not None:
                    return proxied_remote_response
                return RedirectResponse(url=mock_reference)
            if mock_reference.startswith("/mock-files/"):
                mock_reference = mock_reference.split("/mock-files/", 1)[1]
            try:
                return m.get_mock_file(mock_reference)
            except HTTPException as exc:
                if exc.status_code != 404:
                    raise
            resolved_response = resolve_customer_document_reference(mock_reference)
            if resolved_response is not None:
                return resolved_response

        raise HTTPException(status_code=404, detail="File not found")

    @app.get("/api/customer/document/pdf")
    def get_customer_document_pdf(
        recordId: str,
        fieldName: str,
        user: dict[str, Any] | None = Depends(m.get_optional_current_user),
    ) -> Response:
        document_response = get_customer_document(recordId=recordId, fieldName=fieldName, user=user)
        return m._build_pdf_only_response_from_document_response(document_response, fallback_file_name="document.pdf")

    @app.post("/api/customer/search-duplicates")
    def search_duplicate_customers(
        body: dict[str, Any],
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        bot_name = m.normalize_bot_name(user.get("botName") or "Vendor Bot")
        if bot_name == "Vendor Bot":
            raise HTTPException(status_code=403, detail="Access denied for Vendor Bot")

        customer_pan = str(m.empty_to_none(body.get("customerPan")) or "").strip().upper()
        if not customer_pan:
            raise HTTPException(status_code=400, detail="customerPan is required")
        if not m.PAN_REGEX.fullmatch(customer_pan):
            raise HTTPException(status_code=400, detail="Invalid PAN format")

        def normalize_pan(value: Any) -> str:
            return "".join(ch for ch in str(value or "").upper() if ch.isalnum())

        def record_pan(record: dict[str, Any]) -> str:
            return normalize_pan(
                pick_first(
                    record,
                    "PANNumber",
                    "customerPan",
                    "vendorPan",
                    "PANNo",
                    "panNumber",
                    "PANNumber1",
                    "PAN",
                )
            )

        def map_duplicate_row(record: dict[str, Any]) -> dict[str, str]:
            return {
                "customerCode": pick_first(
                    record,
                    "CustomerCode",
                    "SAPCustomerCode",
                    "BusinessPartnerCode",
                    "customerCode",
                    "sapCustomerCode",
                    "VendorCode",
                    "vendorCode",
                ),
                "name": pick_first(
                    record,
                    "CustomerName",
                    "customerName",
                    "vendorName",
                    "VendorName",
                    "TopCustomer",
                    "name",
                ),
                "name1": pick_first(record, "CustomerName1", "Name1", "name1", "VendorName1"),
                "city": pick_first(record, "City", "city", "DistrictCity", "districtCity"),
                "district": pick_first(record, "District", "district"),
                "pan": record_pan(record),
                "gstin": pick_first(record, "GSTIN", "GSTIN1", "GSTNumber", "gstin", "gstNumber"),
            }

        matches: list[dict[str, str]] = []
        seen: set[tuple[str, str, str, str]] = set()
        for source in (
            list(iter_customer_invite_records())
            + list(iter_customer_local_form_records())
            + load_customer_sql_records()
        ):
            if not isinstance(source, dict):
                continue
            if record_pan(source) != customer_pan:
                continue
            row = map_duplicate_row(source)
            dedupe_key = (
                str(row.get("customerCode") or "").strip().lower(),
                str(row.get("name") or "").strip().lower(),
                str(row.get("pan") or "").strip().lower(),
                str(row.get("gstin") or "").strip().lower(),
            )
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            matches.append(row)

        return {
            "customerPan": customer_pan,
            "matchCount": len(matches),
            "matches": matches,
        }

    @app.post("/api/customer/invite")
    def invite_customer(
        body: dict[str, Any],
        request: Request,
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        bot_name = m.normalize_bot_name(user.get("botName") or "Vendor Bot")
        if bot_name == "Vendor Bot":
            raise HTTPException(status_code=403, detail="Access denied for Vendor Bot")

        company_dealing_with = str(m.empty_to_none(body.get("companyDealingWith")) or "").strip()
        customer_name = str(m.empty_to_none(body.get("customerName")) or "").strip()
        customer_email = str(m.empty_to_none(body.get("customerEmail")) or "").strip()
        customer_type_raw = str(m.empty_to_none(body.get("customerType")) or "").strip().lower()
        customer_pan = str(m.empty_to_none(body.get("customerPan")) or "").strip().upper()
        postal_zip_code = str(m.empty_to_none(body.get("postalZipCode")) or "").strip()
        hod_email = str(m.empty_to_none(body.get("hodEmail")) or "").strip()

        if not company_dealing_with or not customer_name or not customer_email or not customer_type_raw:
            raise HTTPException(
                status_code=400,
                detail="companyDealingWith, customerName, customerEmail and customerType are required",
            )
        if "@" not in customer_email or "." not in customer_email.split("@", 1)[-1]:
            raise HTTPException(status_code=400, detail="Invalid customerEmail format")
        if hod_email and ("@" not in hod_email or "." not in hod_email.split("@", 1)[-1]):
            raise HTTPException(status_code=400, detail="Invalid hodEmail format")

        customer_type_map = {"domestic": "Domestic", "import": "Import"}
        customer_type = customer_type_map.get(customer_type_raw)
        if not customer_type:
            raise HTTPException(status_code=400, detail="customerType must be Domestic or Import")
        if customer_type == "Domestic":
            if not customer_pan:
                raise HTTPException(status_code=400, detail="customerPan is required for Domestic customerType")
            if not m.PAN_REGEX.fullmatch(customer_pan):
                raise HTTPException(status_code=400, detail="Invalid customerPan format")
            postal_zip_code = ""
        else:
            customer_pan = ""

        record_id = str(m.uuid.uuid4()).upper()
        created_at = customer_now_ist_iso()
        registration_path = "/web/customer-public-form"
        registration_link = build_customer_public_url(registration_path, request=request)
        invite_status = STATUS_INVITE_CUSTOMER

        invite_record: dict[str, Any] = {
            "recordId": record_id,
            "CompanyDealing": company_dealing_with,
            "vendorCategory": company_dealing_with,
            "vendorName": customer_name,
            "CustomerName": customer_name,
            "vendorEmail": customer_email,
            "rubaminContactPerson": str(user.get("name") or "").strip(),
            "rubaminContactPersonEmail": str(user.get("email") or "").strip(),
            "vendorPan": customer_pan or None,
            "PANNumber": customer_pan or None,
            "postalZipCode": postal_zip_code or None,
            "vendorType": customer_type,
            "CustomerType": customer_type,
            "inviteStatus": invite_status,
            "InviteStatus": invite_status,
            "approverStatus": invite_status,
            "status": invite_status,
            "createdAt": created_at,
            "updatedAt": created_at,
            "hodEmail": hod_email or None,
            "HODEmail": hod_email or None,
        }
        m.invite_forms[record_id] = invite_record

        sql_payload = {
            "recordId": record_id,
            "CompanyDealing": company_dealing_with,
            "vendorCategory": company_dealing_with,
            "CustomerGroup": company_dealing_with,
            "CustomerName": customer_name,
            "vendorName": customer_name,
            "customerEmail": customer_email,
            "vendorEmail": customer_email,
            "rubaminContactPerson": str(user.get("name") or "").strip(),
            "rubaminContactPersonEmail": str(user.get("email") or "").strip(),
            "PANNumber": customer_pan or None,
            "postalZipCode": postal_zip_code or None,
            "PostalZipCode": postal_zip_code or None,
            "inviteStatus": invite_status,
            "InviteStatus": invite_status,
            "approverStatus": invite_status,
            "ApproverStatus": invite_status,
            "ApprovStatus": invite_status,
            "status": invite_status,
            "created_at": created_at,
            "updated_at": created_at,
            "hodEmail": hod_email or None,
            "HODEmail": hod_email or None,
            "vendorType": customer_type,
            "CustomerType": customer_type,
        }

        try:
            customer_db.insert_customer_master_data(sql_payload)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Invite saved in memory but SQL insert failed: {exc}") from exc

        email_sent = False
        email_error: str | None = None
        try:
            m.send_vendor_invite_email(
                vendor_name=customer_name,
                vendor_email=customer_email,
                record_id=record_id,
                registration_url=registration_link,
                cc_email=str(user.get("email") or "").strip() or None,
                entity_label="Customer",
            )
            email_sent = True
            invite_record["inviteEmailSentAt"] = customer_now_ist_iso()
        except Exception as exc:
            email_error = str(exc)

        return {
            "message": "Customer invitation saved",
            "recordId": record_id,
            "emailSent": email_sent,
            "emailError": email_error,
            "registrationLink": registration_link,
        }

    @app.post("/api/customer/public-submit")
    async def submit_customer_public_form(request: Request) -> dict[str, Any]:
        form_data = await request.form()
        record_id = m._normalize_record_id(form_data.get("recordId"))
        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")

        selected_sections_raw = str(form_data.get("selectedSections") or "").strip()
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

        def is_section_selected(*keywords: str) -> bool:
            if not selected_sections:
                return True
            normalized_keywords = [str(keyword or "").strip().lower() for keyword in keywords if str(keyword or "").strip()]
            if not normalized_keywords:
                return True
            for section_name in selected_sections:
                if any(keyword in section_name for keyword in normalized_keywords):
                    return True
            return False

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Invalid or expired recordId")

        updates_local: dict[str, Any] = {}
        updates_sql: dict[str, Any] = {}

        def read_effective_value(field_name: str) -> Any:
            target = str(field_name or "").strip().lower()
            if not target:
                return None
            for source in (updates_sql, updates_local, local_record, sql_record):
                if not isinstance(source, dict):
                    continue
                for key, value in source.items():
                    if str(key or "").strip().lower() != target:
                        continue
                    if value is None:
                        continue
                    if isinstance(value, str):
                        trimmed = value.strip()
                        if trimmed:
                            return trimmed
                    elif isinstance(value, (bytes, bytearray, memoryview)):
                        if len(value) > 0:
                            return value
                    else:
                        return value
            return None

        def has_uploaded_file(field_name: str) -> bool:
            value = form_data.get(field_name)
            if value is None:
                return False
            if isinstance(value, (StarletteUploadFile,)) or (
                hasattr(value, "filename") and hasattr(value, "file")
            ):
                return bool(str(getattr(value, "filename", "") or "").strip())
            return False

        def has_existing_document(field_name: str) -> bool:
            for candidate in (
                field_name,
                f"{field_name}_FileName",
                f"{field_name}FileName",
                f"{field_name}_filename",
                f"{field_name}filename",
            ):
                if read_effective_value(candidate) is not None:
                    return True
            return False

        for key, value in form_data.multi_items():
            if key == "recordId":
                continue
            if is_obsolete_customer_field(key):
                continue
            if isinstance(value, (StarletteUploadFile,)):
                continue
            if hasattr(value, "filename") and hasattr(value, "file"):
                continue
            normalized = m.empty_to_none(str(value)) if isinstance(value, str) else value
            updates_local[key] = normalized
            updates_sql[key] = normalized

        for key, value in form_data.multi_items():
            if is_obsolete_customer_field(key):
                continue
            if not isinstance(value, (StarletteUploadFile,)) and not (
                hasattr(value, "filename") and hasattr(value, "file")
            ):
                continue

            raw_filename = str(getattr(value, "filename", "") or "").strip()
            filename = Path(raw_filename).name
            if not filename:
                continue

            content = await value.read()
            raw_content_type = str(getattr(value, "content_type", "") or "").strip()
            content_type = normalize_inline_document_content_type(filename, raw_content_type)
            file_size = len(content or b"")

            url = m._store_uploaded_document(record_id, key, filename, content, content_type)
            updates_local[key] = url
            updates_local[f"{key}_FileName"] = filename
            updates_local[f"{key}_ContentType"] = content_type
            updates_local[f"{key}_FileSize"] = file_size

            updates_sql[key] = content
            updates_sql[f"{key}_FileName"] = filename
            updates_sql[f"{key}_ContentType"] = content_type
            updates_sql[f"{key}_FileSize"] = file_size

        submitted_validation_snapshot = dict(updates_local)

        gst_status_value = str(read_effective_value("GSTStatus") or "").strip().lower()
        if not is_section_selected("gst"):
            gst_status_value = ""
        gst_validation_status = ""
        gst_validation_error = ""
        if gst_status_value == "registered":
            updates_local["PANStatus"] = "Yes"
            updates_sql["PANStatus"] = "Yes"
            pan_from_gstin = extract_pan_from_gstin(read_effective_value("GSTIN"))
            if pan_from_gstin:
                updates_local["PANNumber"] = pan_from_gstin
                updates_sql["PANNumber"] = pan_from_gstin

            missing_gst_fields: list[str] = []
            if not str(read_effective_value("GSTIN") or "").strip():
                missing_gst_fields.append("GST Number")
            if not str(read_effective_value("TradeNameGST") or "").strip():
                missing_gst_fields.append("Trade Name as per GST")
            if not (has_uploaded_file("GSTCertificate") or has_existing_document("GSTCertificate")):
                missing_gst_fields.append("Upload GST Certificate")
            if missing_gst_fields:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        "Missing mandatory fields for GST Registered status: "
                        + ", ".join(missing_gst_fields)
                    ),
                )

            gst_validation_status = "Not Validated"
            gst_number_for_validation = str(read_effective_value("GSTIN") or "").strip().upper()
            try:
                gst_validation_payload = m._run_gst_extraction_inline(gst_number_for_validation) or {}
                taxpayer_type_from_validation = m._extract_gst_taxpayer_type(gst_validation_payload)
                gst_status_from_validation = normalize_status_key(
                    gst_validation_payload.get("GSTIN / UIN  Status")
                    or gst_validation_payload.get("GSTIN / UIN Status")
                    or gst_validation_payload.get("GST Status")
                    or gst_validation_payload.get("Status")
                )
                if taxpayer_type_from_validation:
                    updates_local["taxpayerTypeGst"] = taxpayer_type_from_validation
                    updates_local["TaxpayerTypeGST"] = taxpayer_type_from_validation
                    updates_sql["TaxpayerTypeGST"] = taxpayer_type_from_validation
                if gst_status_from_validation in {"valid", "active"}:
                    gst_validation_status = "Valid"
                elif gst_status_from_validation:
                    gst_validation_status = "Invalid"
            except Exception as exc:
                gst_validation_error = str(exc)

            updates_local["GST Valid"] = gst_validation_status
            updates_sql["GST Valid"] = gst_validation_status

        gst_change_note = m._build_validation_change_note(
            submitted_validation_snapshot,
            None,
            updates_local,
            m.GST_VALIDATION_CHANGE_FIELDS,
            "GST Details",
        )
        if gst_change_note:
            for note_key in m.GST_VALIDATION_CHANGE_NOTE_KEYS:
                updates_local[note_key] = gst_change_note
                updates_sql[note_key] = gst_change_note

        pan_status_value = str(read_effective_value("PANStatus") or "").strip().lower()
        if not is_section_selected("pan"):
            pan_status_value = ""
        if pan_status_value in {"yes", "available"}:
            missing_pan_fields: list[str] = []
            if not str(read_effective_value("PANNumber") or "").strip():
                missing_pan_fields.append("PAN Number")
            if not str(read_effective_value("TradeNamePAN") or "").strip():
                missing_pan_fields.append("Trade Name as per PAN")
            if not (has_uploaded_file("PANUpload") or has_existing_document("PANUpload")):
                missing_pan_fields.append("Upload PAN Card")
            if missing_pan_fields:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        "Missing mandatory fields for PAN Yes status: "
                        + ", ".join(missing_pan_fields)
                    ),
                )

        company_status_raw = str(
            read_effective_value("CompanyStatus")
            or read_effective_value("CINStatus")
            or read_effective_value("cinStatus")
            or ""
        ).strip()
        company_status_key = company_status_raw.lower()
        if company_status_key in {"yes", "y", "true", "1", "registered", "available"}:
            normalized_company_status = "yes"
        elif company_status_key in {"no", "n", "false", "0", "unregistered", "not available"}:
            normalized_company_status = "no"
        elif company_status_key:
            normalized_company_status = "yes"
        else:
            normalized_company_status = ""

        if not normalized_company_status and is_section_selected("cin", "company status", "pan"):
            raise HTTPException(status_code=400, detail="CIN Status is required")

        if normalized_company_status:
            normalized_company_status_text = "Yes" if normalized_company_status == "yes" else "No"
            updates_local["CompanyStatus"] = normalized_company_status_text
            updates_sql["CompanyStatus"] = normalized_company_status_text

        if normalized_company_status == "yes" and is_section_selected("cin", "company status", "pan"):
            missing_company_fields: list[str] = []
            if not str(read_effective_value("CIN") or "").strip():
                missing_company_fields.append("CIN Number")
            if not (
                has_uploaded_file("IncorporationCertificate")
                or has_existing_document("IncorporationCertificate")
            ):
                missing_company_fields.append("Upload Incorporation Certificate")
            if missing_company_fields:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        "Missing mandatory fields for CIN Status Yes: "
                        + ", ".join(missing_company_fields)
                    ),
                )

        updates_local["approverStatus"] = STATUS_PENDING_USER_APPROVAL
        updates_local["status"] = STATUS_PENDING_USER_APPROVAL
        updates_local["updatedAt"] = customer_now_ist_iso()

        updates_sql["approverStatus"] = STATUS_PENDING_USER_APPROVAL
        updates_sql["ApproverStatus"] = STATUS_PENDING_USER_APPROVAL
        updates_sql["ApprovStatus"] = STATUS_PENDING_USER_APPROVAL
        updates_sql["status"] = STATUS_PENDING_USER_APPROVAL
        updates_sql["updated_at"] = customer_now_ist_iso()

        local_updated = m.update_local_vendor_record_by_record_id(record_id, updates_local)
        sql_updated = customer_db.update_customer_record_by_record_id(record_id, updates_sql)

        if not sql_updated and not sql_record:
            insert_payload = dict(updates_sql)
            insert_payload["recordId"] = record_id
            try:
                customer_db.insert_customer_master_data(insert_payload)
                sql_updated = True
            except Exception:
                sql_updated = False

        if sql_record and not sql_updated:
            raise HTTPException(status_code=500, detail="Customer record found in database but update failed")
        if not local_updated and not sql_updated:
            raise HTTPException(status_code=400, detail="Unable to submit customer form")

        refreshed_local, refreshed_sql = get_customer_records(record_id)
        merged_record = merge_customer_records(record_id, refreshed_local, refreshed_sql)
        customer_name = get_customer_display_name(merged_record, local_record, sql_record)
        submitted_on = str(
            m._first_existing_record_value(
                merged_record,
                None,
                "updatedAt",
                "UpdatedAt",
                "updated_at",
                "createdAt",
                "CreatedAt",
                "created_at",
            )
            or customer_now_ist_iso()
        )
        dashboard_url = build_customer_dashboard_url(request)
        next_status = STATUS_PENDING_USER_APPROVAL
        user_contact_name = (
            m._extract_record_buyer_name(merged_record)
            or m._extract_record_buyer_name(refreshed_local)
            or m._extract_record_buyer_name(refreshed_sql)
        )
        user_contact_email = (
            m._extract_record_buyer_email(merged_record)
            or m._extract_record_buyer_email(refreshed_local)
            or m._extract_record_buyer_email(refreshed_sql)
        )
        approval_email_sent_count, approval_email_error = m._send_stage_approval_notifications(
            stage="user",
            record_id=record_id,
            vendor_name=customer_name,
            submitted_on=submitted_on,
            approver_status=next_status,
            dashboard_url=dashboard_url,
            preferred_contact_name=user_contact_name,
            preferred_contact_email=user_contact_email,
            merged_record=merged_record,
            local_record=refreshed_local,
            db_record=refreshed_sql,
            entity_label="Customer",
        )

        return {
            "message": "Customer form submitted successfully",
            "recordId": record_id,
            "approvalEmailSent": approval_email_sent_count > 0,
            "approvalEmailSentCount": approval_email_sent_count,
            "approvalEmailError": approval_email_error,
            "gstValidationStatus": gst_validation_status or None,
            "gstValidationError": gst_validation_error or None,
            "nextUrl": "/web/login",
        }

    @app.post("/api/customer/internal-submit")
    async def submit_customer_internal_form(request: Request) -> dict[str, Any]:
        form_data = await request.form()
        record_id = str(m.uuid.uuid4()).upper()
        try:
            session_user = m.get_cookie_session_user(request)
        except HTTPException as exc:
            raise HTTPException(status_code=401, detail="Login required for internal customer form submission") from exc

        actor_email = str(session_user.get("email") or "").strip()

        updates_local: dict[str, Any] = {}
        updates_sql: dict[str, Any] = {}

        for key, value in form_data.multi_items():
            if key == "recordId":
                continue
            if is_obsolete_customer_field(key):
                continue
            if isinstance(value, (StarletteUploadFile,)):
                continue
            if hasattr(value, "filename") and hasattr(value, "file"):
                continue
            normalized = m.empty_to_none(str(value)) if isinstance(value, str) else value
            updates_local[key] = normalized
            updates_sql[key] = normalized

        for key, value in form_data.multi_items():
            if is_obsolete_customer_field(key):
                continue
            if not isinstance(value, (StarletteUploadFile,)) and not (
                hasattr(value, "filename") and hasattr(value, "file")
            ):
                continue

            raw_filename = str(getattr(value, "filename", "") or "").strip()
            filename = Path(raw_filename).name
            if not filename:
                continue

            content = await value.read()
            raw_content_type = str(getattr(value, "content_type", "") or "").strip()
            content_type = normalize_inline_document_content_type(filename, raw_content_type)
            file_size = len(content or b"")

            url = m._store_uploaded_document(record_id, key, filename, content, content_type)
            updates_local[key] = url
            updates_local[f"{key}_FileName"] = filename
            updates_local[f"{key}_ContentType"] = content_type
            updates_local[f"{key}_FileSize"] = file_size

            updates_sql[key] = content
            updates_sql[f"{key}_FileName"] = filename
            updates_sql[f"{key}_ContentType"] = content_type
            updates_sql[f"{key}_FileSize"] = file_size

        now_value = customer_now_ist_iso()
        target_status = STATUS_PENDING_VALIDATOR_APPROVAL
        submitted_validation_snapshot = dict(updates_local)

        # Run GST/MSME extraction for internal submit based on selected status values (best effort).
        try:
            should_enrich = str(target_status or "").strip().lower() in {
                "pending for user approval",
                "pendign for user approval",
                "pending for validator approval",
                "pendign for validator approval",
                "pending for validation approval",
            }

            def read_submitted_text(*keys: str) -> str:
                for key_name in keys:
                    value = updates_local.get(key_name)
                    if value is None:
                        value = updates_sql.get(key_name)
                    if value is None:
                        continue
                    text = str(value).strip()
                    if text:
                        return text
                return ""

            gst_status_raw = read_submitted_text(
                "GSTStatus",
                "gstStatus",
                "gstRegistrationStatus",
                "GST Registration Status",
            ).lower()
            gst_number_raw = read_submitted_text(
                "GSTIN",
                "gstNumber",
                "gstin",
                "GST Number",
            ).upper()
            should_run_gst = (
                should_enrich
                and gst_number_raw
                and (
                    "registered" in gst_status_raw
                    or gst_status_raw in {"yes", "y", "true", "1"}
                )
            )
            if should_run_gst:
                print(f"[GST][CUSTOMER INTERNAL] Running inline extraction for GSTIN={gst_number_raw!r}")
                gst_data = m._run_gst_extraction_inline(gst_number_raw) or {}
                trade_name = str(gst_data.get("Trade Name") or "").strip()
                taxpayer_type = m._extract_gst_taxpayer_type(gst_data)
                gst_validation_status = "Not Validated"
                gst_status_from_validation = str(gst_data.get("GSTIN / UIN  Status") or "").strip().lower()
                if gst_status_from_validation in {"valid", "active"}:
                    gst_validation_status = "Valid"
                elif gst_status_from_validation:
                    gst_validation_status = "Invalid"

                updates_local["GST Valid"] = gst_validation_status
                updates_sql["GST Valid"] = gst_validation_status
                if trade_name:
                    updates_local["TradeNameGST"] = trade_name
                    updates_local["TradeName"] = trade_name
                    updates_sql["TradeNameGST"] = trade_name
                    updates_sql["TradeName"] = trade_name
                if taxpayer_type:
                    updates_local["TaxpayerTypeGST"] = taxpayer_type
                    updates_local["taxpayerTypeGst"] = taxpayer_type
                    updates_sql["TaxpayerTypeGST"] = taxpayer_type
                    updates_sql["taxpayerTypeGst"] = taxpayer_type

            msme_status_raw = read_submitted_text(
                "MSMEStatus",
                "msmeStatus",
                "MSME Status",
            ).lower()
            udyam_number_raw = read_submitted_text(
                "UdyamNumber",
                "udyamNumber",
                "Udyam Number",
            ).upper()
            should_run_msme = (
                should_enrich
                and udyam_number_raw
                and (
                    "registered" in msme_status_raw
                    or msme_status_raw in {"yes", "y", "true", "1"}
                )
            )
            if should_run_msme:
                print(f"[MSME][CUSTOMER INTERNAL] Running inline extraction for Udyam={udyam_number_raw!r}")
                msme_data = m._run_msme_extraction_inline(udyam_number_raw) or {}
                msme_category = str(msme_data.get("MSME Category") or "").strip()
                msme_industry = str(msme_data.get("MSME Industry") or "").strip()
                udyam_reg_date = str(msme_data.get("Date of Udyam Registration") or "").strip()

                if msme_category:
                    updates_local["MSMECategory"] = msme_category
                    updates_local["msmeCategory"] = msme_category
                    updates_sql["MSMECategory"] = msme_category
                    updates_sql["msmeCategory"] = msme_category
                if msme_industry:
                    updates_local["MSMEIndustry"] = msme_industry
                    updates_local["msmeIndustry"] = msme_industry
                    updates_sql["MSMEIndustry"] = msme_industry
                    updates_sql["msmeIndustry"] = msme_industry
                if udyam_reg_date:
                    updates_local["UdyamRegistrationDate"] = udyam_reg_date
                    updates_local["udyamRegistrationDate"] = udyam_reg_date
                    updates_sql["UdyamRegistrationDate"] = udyam_reg_date
                    updates_sql["udyamRegistrationDate"] = udyam_reg_date
        except Exception as exc:
            # Enrichment is best effort; form submit should not fail.
            print(f"[ENRICHMENT][CUSTOMER INTERNAL] GST/MSME enrichment skipped due to error: {exc}")

        gst_change_note = m._build_validation_change_note(
            submitted_validation_snapshot,
            None,
            updates_local,
            m.GST_VALIDATION_CHANGE_FIELDS,
            "GST Details",
        )
        if gst_change_note:
            for note_key in m.GST_VALIDATION_CHANGE_NOTE_KEYS:
                updates_local[note_key] = gst_change_note
                updates_sql[note_key] = gst_change_note

        msme_change_note = m._build_validation_change_note(
            submitted_validation_snapshot,
            None,
            updates_local,
            m.MSME_VALIDATION_CHANGE_FIELDS,
            "MSME Details",
        )
        if msme_change_note:
            for note_key in m.MSME_VALIDATION_CHANGE_NOTE_KEYS:
                updates_local[note_key] = msme_change_note
                updates_sql[note_key] = msme_change_note

        updates_local["recordId"] = record_id
        updates_local["approverStatus"] = target_status
        updates_local["status"] = target_status
        updates_local["createdAt"] = now_value
        updates_local["updatedAt"] = now_value

        updates_sql["recordId"] = record_id
        updates_sql["approverStatus"] = target_status
        updates_sql["ApproverStatus"] = target_status
        updates_sql["ApprovStatus"] = target_status
        updates_sql["status"] = target_status
        updates_sql["created_at"] = now_value
        updates_sql["updated_at"] = now_value

        if actor_email:
            actor_name = str(session_user.get("name") or "").strip()
            updates_local["rubaminContactPersonEmail"] = actor_email
            updates_sql["rubaminContactPersonEmail"] = actor_email
            if actor_name:
                updates_local["rubaminContactPerson"] = actor_name
                updates_sql["rubaminContactPerson"] = actor_name

        try:
            inserted_record_id = customer_db.insert_customer_master_data(dict(updates_sql))
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Unable to submit customer form: {exc}") from exc

        persisted_record_id = m._normalize_record_id(inserted_record_id) or record_id
        updates_local["recordId"] = persisted_record_id
        m.update_local_vendor_record_by_record_id(persisted_record_id, updates_local)

        refreshed_local, refreshed_sql = get_customer_records(persisted_record_id)
        merged_record = merge_customer_records(persisted_record_id, refreshed_local, refreshed_sql)
        customer_name = get_customer_display_name(merged_record, refreshed_local, refreshed_sql)
        submitted_on = str(
            m._first_existing_record_value(
                merged_record,
                None,
                "updatedAt",
                "UpdatedAt",
                "updated_at",
                "createdAt",
                "CreatedAt",
                "created_at",
            )
            or now_value
        )
        dashboard_url = build_customer_dashboard_url(request)
        approval_email_sent_count, approval_email_error = m._send_stage_approval_notifications(
            stage="validator",
            record_id=persisted_record_id,
            vendor_name=customer_name,
            submitted_on=submitted_on,
            approver_status=target_status,
            dashboard_url=dashboard_url,
            merged_record=merged_record,
            local_record=refreshed_local,
            db_record=refreshed_sql,
            entity_label="Customer",
        )

        return {
            "message": "Customer form submitted successfully",
            "recordId": persisted_record_id,
            "approvalEmailSent": approval_email_sent_count > 0,
            "approvalEmailSentCount": approval_email_sent_count,
            "approvalEmailError": approval_email_error,
            "nextUrl": "/web/customer-dashboard",
        }

    @app.post("/api/customer/cancel")
    def cancel_customer_record(
        body: dict[str, Any],
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        record_id = m._normalize_record_id(m.empty_to_none(body.get("recordId")))
        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Customer record not found")

        normalized_role = m.normalize_role(user.get("role"))
        if normalized_role == "validator":
            raise HTTPException(status_code=403, detail="Validator role cannot cancel customer")
        if normalized_role != "admin":
            user_email = str(user.get("email") or "").strip().lower()
            allowed_emails = m._collect_record_access_emails(local_record, sql_record)
            if not user_email or not allowed_emails or user_email not in allowed_emails:
                raise HTTPException(status_code=403, detail="Access denied for this record")

        target_status = "Customer Cancelled"
        local_updated = m.update_local_vendor_status_by_record_id(record_id, target_status)
        sql_updated = customer_db.update_customer_approver_status_by_record_id(record_id, target_status)
        if not local_updated and not sql_updated:
            raise HTTPException(status_code=404, detail="Customer record not found")

        return {
            "message": "Customer cancelled successfully",
            "recordId": record_id,
            "approverStatus": target_status,
        }

    @app.put("/api/customer/update-record")
    def update_customer_record(
        body: dict[str, Any],
        request: Request,
        user: dict[str, Any] | None = Depends(m.get_optional_current_user),
    ) -> dict[str, Any]:
        record_id = m._normalize_record_id(m.empty_to_none(body.get("recordId")))
        updates_raw = body.get("updates")

        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")
        if not isinstance(updates_raw, dict) or not updates_raw:
            raise HTTPException(status_code=400, detail="updates are required")

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Customer record not found")

        current_status = str(
            m._first_existing_record_value(
                sql_record,
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
        resubmission_stage, resubmission_target_status = m._resolve_resubmission_target_from_status(current_status)

        if user is not None:
            normalized_role = m.normalize_role(user.get("role"))
            if normalized_role == "validator":
                raise HTTPException(status_code=403, detail="Validator role cannot use update-record endpoint")
            if normalized_role != "admin":
                user_email = str(user.get("email") or "").strip().lower()
                allowed_emails = m._collect_record_access_emails(local_record, sql_record)
                if not user_email or not allowed_emails or user_email not in allowed_emails:
                    raise HTTPException(status_code=403, detail="Access denied for this record")

        merged_record_for_lock = merge_customer_records(record_id, local_record, sql_record)
        locked_update_keys = m.get_validator_locked_update_keys(merged_record_for_lock)

        protected_keys = {"recordid", "record", "id"}
        updates: dict[str, Any] = {}
        for key, value in updates_raw.items():
            key_name = str(key or "").strip()
            if not key_name:
                continue
            if is_obsolete_customer_field(key_name):
                continue
            key_name_lower = key_name.lower()
            if key_name_lower in protected_keys or key_name_lower in locked_update_keys:
                continue
            updates[key_name] = m.empty_to_none(value) if isinstance(value, str) else value

        updates = sanitize_customer_updates(updates)

        if resubmission_target_status:
            updates["approverStatus"] = resubmission_target_status
            updates["ApproverStatus"] = resubmission_target_status
            updates["ApprovStatus"] = resubmission_target_status

        if not updates:
            raise HTTPException(status_code=400, detail="No updatable fields found")

        db_updated = customer_db.update_customer_record_by_record_id(record_id, updates)
        if sql_record and not db_updated:
            raise HTTPException(status_code=500, detail="Customer record found in database but update failed")

        local_updated = m.update_local_vendor_record_by_record_id(record_id, updates)
        if not local_updated and not db_updated:
            raise HTTPException(status_code=400, detail="Unable to update customer record")

        refreshed_local, refreshed_sql = get_customer_records(record_id)
        merged_record = merge_customer_records(record_id, refreshed_local, refreshed_sql)

        approval_email_sent_count = 0
        approval_email_error: str | None = None
        if resubmission_target_status and resubmission_stage:
            dashboard_url = build_customer_dashboard_url(request)
            update_remarks_for_mail = str(
                m.empty_to_none(
                    updates.get("updateRemarks")
                    or updates.get("UpdateRemarks")
                    or updates.get("remarks")
                    or updates.get("Remarks")
                )
                or ""
            ).strip()
            customer_name = get_customer_display_name(merged_record, local_record, sql_record)
            submitted_on = str(
                m._first_existing_record_value(
                    merged_record,
                    None,
                    "updatedAt",
                    "UpdatedAt",
                    "updated_at",
                    "createdAt",
                    "CreatedAt",
                    "created_at",
                )
                or customer_now_ist_iso()
            )
            approval_email_sent_count, approval_email_error = m._send_stage_approval_notifications(
                stage=resubmission_stage,
                record_id=record_id,
                vendor_name=customer_name,
                submitted_on=submitted_on,
                approver_status=resubmission_target_status,
                dashboard_url=dashboard_url,
                remarks=update_remarks_for_mail,
                merged_record=merged_record,
                local_record=local_record,
                db_record=sql_record,
                entity_label="Customer",
            )

        return {
            "message": (
                "Customer record resubmitted successfully"
                if resubmission_target_status
                else "Customer record updated successfully"
            ),
            "recordId": record_id,
            "approverStatus": (
                str(resubmission_target_status)
                if resubmission_target_status
                else str(
                    m._first_existing_record_value(
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
            ),
            "approvalEmailSent": approval_email_sent_count > 0,
            "approvalEmailSentCount": approval_email_sent_count,
            "approvalEmailError": approval_email_error,
            "resubmissionStage": resubmission_stage,
            "vendor": merged_record,
        }

    @app.put("/api/customer/update-record-with-files")
    async def update_customer_record_with_files(
        request: Request,
        user: dict[str, Any] | None = Depends(m.get_optional_current_user),
    ) -> dict[str, Any]:
        def uses_filename_value(key_name: str) -> bool:
            key_lower = str(key_name or "").strip().lower()
            return (
                key_lower.endswith("filename")
                or key_lower.endswith("_filename")
                or key_lower.endswith("_name")
            )

        form_data = await request.form()
        record_id = m._normalize_record_id(m.empty_to_none(form_data.get("recordId")))
        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")

        updates_raw_value = form_data.get("updates")
        updates_payload: dict[str, Any] = {}
        if isinstance(updates_raw_value, str) and updates_raw_value.strip():
            try:
                parsed_updates = m.json.loads(updates_raw_value)
            except m.json.JSONDecodeError:
                raise HTTPException(status_code=400, detail="updates must be valid JSON")
            if not isinstance(parsed_updates, dict):
                raise HTTPException(status_code=400, detail="updates must be a JSON object")
            updates_payload = sanitize_customer_updates(dict(parsed_updates))

        doc_source_map: dict[str, str] = {}
        doc_source_map_raw = form_data.get("docSourceMap")
        if isinstance(doc_source_map_raw, str) and doc_source_map_raw.strip():
            try:
                parsed_doc_source = m.json.loads(doc_source_map_raw)
            except m.json.JSONDecodeError:
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
            if not isinstance(value, (StarletteUploadFile,)):
                continue
            key_name = str(key or "").strip()
            if not key_name.startswith("file_"):
                continue

            field_name = key_name[len("file_"):].strip()
            if not field_name:
                continue
            if is_obsolete_customer_field(field_name):
                continue

            raw_filename = str(value.filename or "").strip()
            normalized_filename = Path(raw_filename).name
            if not normalized_filename:
                continue

            file_bytes = await value.read()
            content = file_bytes or b""
            content_type = normalize_inline_document_content_type(
                normalized_filename,
                value.content_type,
            )
            doc_url = m._store_uploaded_document(record_id, field_name, normalized_filename, content, content_type)
            uploaded_doc_updates[field_name] = {
                "url": doc_url,
                "filename": normalized_filename,
            }
            db_file_updates[field_name] = content
            db_file_updates[f"{field_name}_FileName"] = normalized_filename
            db_file_updates[f"{field_name}_ContentType"] = content_type
            db_file_updates[f"{field_name}_FileSize"] = len(content)

        for field_name, file_meta in uploaded_doc_updates.items():
            if is_obsolete_customer_field(field_name):
                continue
            doc_url = str(file_meta.get("url") or "")
            filename = str(file_meta.get("filename") or "")

            target_keys: set[str] = {field_name}
            mapped_source_key = str(doc_source_map.get(field_name) or "").strip()
            if mapped_source_key:
                target_keys.add(mapped_source_key)
            for alias_key in m.DOCUMENT_FIELD_ALIASES.get(field_name, []):
                alias_name = str(alias_key or "").strip()
                if alias_name:
                    target_keys.add(alias_name)

            for key_name in target_keys:
                if is_obsolete_customer_field(key_name):
                    continue
                updates_payload[key_name] = filename if uses_filename_value(key_name) else doc_url

        updates_payload = sanitize_customer_updates(updates_payload)
        db_file_updates = sanitize_customer_updates(db_file_updates)

        if not updates_payload:
            raise HTTPException(status_code=400, detail="updates are required")

        result = update_customer_record({"recordId": record_id, "updates": updates_payload}, request=request, user=user)
        if db_file_updates:
            customer_db.update_customer_record_by_record_id(record_id, db_file_updates)
        return result

    @app.post("/api/customer/review-decision")
    def submit_customer_review_decision(
        body: dict[str, Any],
        request: Request,
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        record_id = m._normalize_record_id(m.empty_to_none(body.get("recordId")))
        decision = str(m.empty_to_none(body.get("decision")) or "").strip().lower()
        remarks = str(m.empty_to_none(body.get("remarks")) or "").strip()
        team_raw = body.get("team")
        validations_raw = body.get("validations")
        tax_details_raw = body.get("taxDetails")
        field_updates_raw = body.get("fieldUpdates")
        approval_mode = str(m.empty_to_none(body.get("approvalMode")) or "").strip().lower()
        reject_to = str(m.empty_to_none(body.get("rejectTo")) or "").strip()
        reject_target = normalize_customer_reject_target(reject_to)

        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")
        if decision not in {"approve", "reject"}:
            raise HTTPException(status_code=400, detail="decision must be Approve or Reject")
        if not remarks:
            raise HTTPException(status_code=400, detail="remarks are required")

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Customer record not found")

        base_role = m.normalize_role(user.get("role"))
        user_email = str(user.get("email") or "").strip().lower()
        user_email_raw = str(user.get("email") or "").strip()
        bot_name = m.normalize_bot_name(user.get("botName") or "Vendor Bot")
        if base_role in {"vendor", "customer"} and bot_name in {"Both", "Customer Bot"}:
            base_role = "user"

        pending_validator_statuses = {
            "pending for validator approval",
            "pendign for validator approval",
        }
        pending_hod_statuses = {
            "pending for hod approval",
            "pendign for hod approval",
        }
        current_status = str(
            m._first_existing_record_value(
                sql_record,
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

        current_status_key = normalize_status_key(current_status)

        def resolve_customer_decision_role(role_value: str, mode_value: str, status_key: str) -> str:
            role_key = m.normalize_role(role_value)
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
                if status_key in {"pending for user approval", "pendign for user approval", "pending for buyer approval", "pendign for buyer approval", "pending for data validation"}:
                    return "user"
                return "validator"

            if role_key == "validator_hod":
                if mode_stage in {"validator", "hod"}:
                    return mode_stage
                if status_key in {"pending for hod approval", "pendign for hod approval"}:
                    return "hod"
                return "validator"

            return role_key

        normalized_role = resolve_customer_decision_role(base_role, approval_mode, current_status_key)

        mode_to_stage = {
            "validator": "validator",
            "hod": "hod",
            "buyer": "user",
            "user": "user",
        }
        decision_stage = mode_to_stage.get(approval_mode)
        if not decision_stage:
            if normalized_role in {"validator", "hod"}:
                decision_stage = normalized_role
            elif normalized_role in {"user", "buyer"}:
                decision_stage = "user"
            elif normalized_role == "admin":
                if is_pending_validator_status(current_status):
                    decision_stage = "validator"
                elif is_pending_hod_status(current_status):
                    decision_stage = "hod"
                else:
                    decision_stage = "user"
            else:
                decision_stage = "user"

        if normalized_role == "hod" and decision_stage == "hod":
            allowed_hod_emails = {
                email
                for email in (
                    m._extract_record_hod_email(local_record),
                    m._extract_record_hod_email(sql_record),
                )
                if email
            }
            if not user_email or not allowed_hod_emails or user_email not in allowed_hod_emails:
                raise HTTPException(status_code=403, detail="Access denied for this record")
        elif normalized_role not in {"admin", "validator", "hod"}:
            allowed_emails = m._collect_record_access_emails(local_record, sql_record)
            if not user_email or not allowed_emails or user_email not in allowed_emails:
                raise HTTPException(status_code=403, detail="Access denied for this record")

        if decision_stage == "hod" and normalized_role not in {"admin", "hod"}:
            raise HTTPException(status_code=403, detail="Only HOD/Admin can submit HOD stage decision")
        if decision_stage == "validator" and normalized_role not in {"admin", "validator", "hod"}:
            raise HTTPException(status_code=403, detail="Only Validator/HOD/Admin can submit Validator stage decision")
        if decision_stage == "user" and normalized_role not in {"admin", "user", "buyer", "vendor", "customer"}:
            raise HTTPException(
                status_code=403,
                detail="Only User/Buyer/Admin (including legacy Vendor/Customer user-stage tokens) can submit User stage decision",
            )

        if decision_stage == "validator" and not is_pending_validator_status(current_status):
            raise HTTPException(
                status_code=409,
                detail=f"Record status does not match Validator stage. Current status: {current_status or 'N/A'}",
            )
        if decision_stage == "hod" and not is_pending_hod_status(current_status):
            raise HTTPException(
                status_code=409,
                detail=f"Record status does not match HOD stage. Current status: {current_status or 'N/A'}",
            )
        if decision_stage == "user" and not is_pending_user_status(current_status):
            raise HTTPException(
                status_code=409,
                detail=f"Record status does not match User stage. Current status: {current_status or 'N/A'}",
            )

        user_role = str(user.get("role") or "User").strip() or "User"
        decision_label = "Approve" if decision == "approve" else "Reject"
        updates: dict[str, Any] = {}
        protected_decision_fields = {
            "recordid",
            "record",
            "id",
            "status",
            "approverstatus",
            "approvstatus",
            "invitestatus",
            "reviewremarks",
            "reviewedbyrole",
            "remarks",
            "reviewdecision",
            "updatedat",
            "updated_at",
            "createdat",
            "created_at",
        }

        def parse_string_value(value: Any) -> str | None:
            if value is None:
                return None
            if isinstance(value, bool):
                return "True" if value else "False"
            if isinstance(value, (int, float)):
                text = str(value).strip()
                return text or None
            if isinstance(value, (list, tuple)):
                for item in value:
                    item_text = str(item or "").strip()
                    if item_text:
                        return item_text
                return None
            if isinstance(value, dict):
                text = str(value).strip()
                return text or None

            text = str(value).strip()
            if not text:
                return None
            try:
                parsed = m.json.loads(text)
                if isinstance(parsed, list):
                    for item in parsed:
                        item_text = str(item or "").strip()
                        if item_text:
                            return item_text
                    return None
                if parsed is None:
                    return None
                if isinstance(parsed, (str, int, float, bool)):
                    parsed_text = str(parsed).strip()
                    return parsed_text or None
            except Exception:
                pass
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

        if isinstance(field_updates_raw, dict):
            for key, value in field_updates_raw.items():
                key_name = str(key or "").strip()
                if not key_name:
                    continue
                key_lower = key_name.lower()
                if key_lower in protected_decision_fields:
                    continue
                if is_obsolete_customer_field(key_name):
                    continue
                updates[key_name] = m.empty_to_none(value) if isinstance(value, str) else value
        stage_role_label = {
            "validator": "Validator",
            "hod": "HOD",
            "user": user_role,
        }.get(decision_stage, user_role)

        if decision_stage == "validator":
            if decision == "approve":
                approver_status = STATUS_PENDING_HOD_APPROVAL
            else:
                approver_status = "Rejected by Validator"
                if reject_target not in {"customer", "buyer"}:
                    raise HTTPException(status_code=400, detail="rejectTo must be Customer or Buyer")

            validations_payload = validations_raw if isinstance(validations_raw, dict) else {}
            gst_validate = parse_bool(validations_payload.get("gstValidate"))
            gst_validate_text = "True" if gst_validate else "False"
            tax_payload = tax_details_raw if isinstance(tax_details_raw, dict) else {}
            tcs_rate_raw = tax_payload.get("tcsRate")
            if tcs_rate_raw is None:
                tcs_rate_raw = body.get("tcsRate")
            turnover_limit_raw = tax_payload.get("turnoverLimit")
            if turnover_limit_raw is None:
                turnover_limit_raw = body.get("turnoverLimit")
            tcs_rate_value = parse_string_value(tcs_rate_raw)
            turnover_limit_value = parse_string_value(turnover_limit_raw)

            updates.update(
                {
                    "approverStatus": approver_status,
                    "ApproverStatus": approver_status,
                    "ApprovStatus": approver_status,
                    "reviewRemarks": remarks,
                    "ReviewRemarks": remarks,
                    "reviewedByRole": stage_role_label,
                    "ReviewedByRole": stage_role_label,
                    "remarks": remarks,
                    "Remarks": remarks,
                    "gstValidate": gst_validate_text,
                    "GST Validate": gst_validate_text,
                    "tcsRate": tcs_rate_value,
                    "TCSRate": tcs_rate_value,
                    "TCS Rate": tcs_rate_value,
                    "turnoverLimit": turnover_limit_value,
                    "TurnoverLimit": turnover_limit_value,
                    "Turnover Limit": turnover_limit_value,
                }
            )
        elif decision_stage == "hod":
            if decision == "approve":
                approver_status = STATUS_PENDING_SAP_CODE_CREATION
            else:
                approver_status = "Rejected by HOD"
                if reject_target not in {"customer", "buyer"}:
                    raise HTTPException(status_code=400, detail="rejectTo must be Customer or Buyer")

            updates.update(
                {
                    "approverStatus": approver_status,
                    "ApproverStatus": approver_status,
                    "ApprovStatus": approver_status,
                    "reviewRemarks": remarks,
                    "ReviewRemarks": remarks,
                    "reviewedByRole": stage_role_label,
                    "ReviewedByRole": stage_role_label,
                    "remarks": remarks,
                    "Remarks": remarks,
                }
            )
        else:
            approver_status = (
                STATUS_PENDING_VALIDATOR_APPROVAL
                if decision == "approve"
                else f"Rejected by {user_role}"
            )
            updates.update(
                {
                    "approverStatus": approver_status,
                    "ApproverStatus": approver_status,
                    "ApprovStatus": approver_status,
                    "reviewRemarks": remarks,
                    "ReviewRemarks": remarks,
                    "reviewedByRole": stage_role_label,
                    "ReviewedByRole": stage_role_label,
                    "remarks": remarks,
                    "Remarks": remarks,
                }
            )

            team_payload = team_raw if isinstance(team_raw, dict) else {}
            approve_required_team_fields = (
                "customerGroup",
                "industryType",
                "incoTerms",
                "paymentTerms",
                "salesOrganization",
                "distributionChannel",
                "division",
                "accountAssignment",
                "currency",
                "scrapSales",
                "insuranceLimit",
                "creditLimit",
                "vendorType",
                "rubaminApproverHod",
            )
            if decision == "approve":
                missing_fields = []
                for field_name in approve_required_team_fields:
                    field_value = (
                        m.empty_to_none(team_payload.get(field_name))
                        if isinstance(team_payload.get(field_name), str)
                        else team_payload.get(field_name)
                    )
                    if field_value is None or (isinstance(field_value, str) and field_value.strip() == ""):
                        missing_fields.append(field_name)
                if missing_fields:
                    raise HTTPException(
                        status_code=400,
                        detail=f"Missing required fields for approve: {', '.join(missing_fields)}",
                    )

            team_field_keys = {
                "customerGroup": ["CustomerGroup"],
                "industryType": ["IndustryType"],
                "incoTerms": ["IncoTerms", "Incoterms"],
                "paymentTerms": ["PaymentTerms"],
                "salesOrganization": ["SalesOrganization"],
                "distributionChannel": ["DistributionChannel"],
                "division": ["Division"],
                "accountAssignment": ["AccountAssignment"],
                "currency": ["Currency"],
                "scrapSales": ["ScrapSales"],
                "insuranceGroup": ["InsuranceGroup"],
                "insuranceLimit": ["InsuranceLimit"],
                "creditLimit": ["CreditLimit"],
                "vendorType": ["VendorType", "CustomerType"],
                "rubaminApproverHod": ["hodEmail", "HodEmail", "HODEmail"],
            }
            for canonical_key, alias_keys in team_field_keys.items():
                if canonical_key not in team_payload:
                    continue
                raw_value = team_payload.get(canonical_key)
                normalized_value = m.empty_to_none(raw_value) if isinstance(raw_value, str) else raw_value
                if canonical_key != "rubaminApproverHod":
                    updates[canonical_key] = normalized_value
                for alias_key in alias_keys:
                    updates[alias_key] = normalized_value

        db_updated = customer_db.update_customer_record_by_record_id(record_id, updates)
        if sql_record and not db_updated:
            raise HTTPException(status_code=500, detail="Customer record found in database but update failed")
        local_updated = m.update_local_vendor_record_by_record_id(record_id, updates)
        if not local_updated and not db_updated:
            raise HTTPException(status_code=400, detail="Unable to update customer review decision")

        refreshed_local, refreshed_sql = get_customer_records(record_id)
        merged_record = merge_customer_records(record_id, refreshed_local, refreshed_sql)

        dashboard_url = build_customer_dashboard_url(request)
        customer_name = get_customer_display_name(merged_record, local_record, sql_record)
        submitted_on = str(
            m._first_existing_record_value(
                merged_record,
                None,
                "updatedAt",
                "UpdatedAt",
                "updated_at",
                "createdAt",
                "CreatedAt",
                "created_at",
            )
            or customer_now_ist_iso()
        )
        update_customer_form_url = build_customer_public_url("/web/update-customer", request=request)
        customer_email = (
            m._extract_record_vendor_email(merged_record)
            or m._extract_record_vendor_email(local_record)
            or m._extract_record_vendor_email(sql_record)
        )
        customer_name_for_mail = (
            get_customer_display_name(merged_record, local_record, sql_record)
            or customer_name
        )
        buyer_email = (
            m._extract_record_buyer_email(merged_record)
            or m._extract_record_buyer_email(local_record)
            or m._extract_record_buyer_email(sql_record)
        )
        buyer_name = (
            m._extract_record_buyer_name(merged_record)
            or m._extract_record_buyer_name(local_record)
            or m._extract_record_buyer_name(sql_record)
        )

        validator_email_sent_count = 0
        validator_email_error: str | None = None
        if (
            decision == "approve"
            and decision_stage == "user"
            and str(approver_status or "").strip().lower() in pending_validator_statuses
        ):
            validator_email_sent_count, validator_email_error = m._send_stage_approval_notifications(
                stage="validator",
                record_id=record_id,
                vendor_name=customer_name,
                submitted_on=submitted_on,
                approver_status=approver_status,
                dashboard_url=dashboard_url,
                remarks=remarks,
                merged_record=merged_record,
                local_record=local_record,
                db_record=sql_record,
                entity_label="Customer",
            )

        hod_approval_email_sent = False
        hod_approval_email_error: str | None = None
        if (
            decision == "approve"
            and decision_stage == "validator"
            and str(approver_status or "").strip().lower() in pending_hod_statuses
        ):
            hod_sent_count, hod_approval_email_error = m._send_stage_approval_notifications(
                stage="hod",
                record_id=record_id,
                vendor_name=customer_name,
                submitted_on=submitted_on,
                approver_status=approver_status,
                dashboard_url=dashboard_url,
                remarks=remarks,
                merged_record=merged_record,
                local_record=local_record,
                db_record=sql_record,
                entity_label="Customer",
            )
            hod_approval_email_sent = hod_sent_count > 0

        rejection_email_sent = False
        rejection_email_error: str | None = None
        if decision == "reject":
            try:
                if decision_stage == "validator":
                    if reject_target == "customer":
                        if not customer_email:
                            raise RuntimeError("Customer email not found")
                        m.send_vendor_rejection_email(
                            recipient_name=customer_name_for_mail,
                            recipient_email=customer_email,
                            record_id=record_id,
                            remarks=remarks,
                            update_form_url=update_customer_form_url,
                            rejected_by="validator",
                            cc_email=buyer_email or None,
                            entity_label="Customer",
                        )
                        rejection_email_sent = True
                    elif reject_target == "buyer":
                        if not buyer_email:
                            raise RuntimeError("Buyer email not found")
                        m.send_vendor_rejection_email(
                            recipient_name=buyer_name,
                            recipient_email=buyer_email,
                            record_id=record_id,
                            remarks=remarks,
                            update_form_url=update_customer_form_url,
                            rejected_by="validator",
                            entity_label="Customer",
                        )
                        rejection_email_sent = True
                elif decision_stage == "hod":
                    if reject_target == "customer":
                        if not customer_email:
                            raise RuntimeError("Customer email not found")
                        m.send_vendor_rejection_email(
                            recipient_name=customer_name_for_mail,
                            recipient_email=customer_email,
                            record_id=record_id,
                            remarks=remarks,
                            update_form_url=update_customer_form_url,
                            rejected_by="HOD",
                            cc_email=buyer_email or None,
                            entity_label="Customer",
                        )
                        rejection_email_sent = True
                    elif reject_target == "buyer":
                        if not buyer_email:
                            raise RuntimeError("Buyer email not found")
                        m.send_vendor_rejection_email(
                            recipient_name=buyer_name,
                            recipient_email=buyer_email,
                            record_id=record_id,
                            remarks=remarks,
                            update_form_url=update_customer_form_url,
                            rejected_by="HOD",
                            entity_label="Customer",
                        )
                        rejection_email_sent = True
                else:
                    if not customer_email:
                        raise RuntimeError("Customer email not found")
                    cc_buyer = buyer_email or user_email_raw
                    m.send_vendor_rejection_email(
                        recipient_name=customer_name_for_mail,
                        recipient_email=customer_email,
                        record_id=record_id,
                        remarks=remarks,
                        update_form_url=update_customer_form_url,
                        rejected_by="buyer" if normalized_role == "buyer" else "user",
                        cc_email=cc_buyer or None,
                        entity_label="Customer",
                    )
                    rejection_email_sent = True
            except Exception as exc:
                rejection_email_error = str(exc)

        return {
            "message": "Customer review decision submitted successfully",
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
            "sapSyncStarted": False,
            "sapSyncQueuedCount": 0,
            "vendor": merged_record,
        }

    @app.post("/api/customer/update-hod-email")
    def update_customer_hod_email(
        body: dict[str, Any],
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        if not m.user_has_any_role(user, "user", "buyer"):
            raise HTTPException(status_code=403, detail="Only User role can update HOD email")

        record_id = m._normalize_record_id(m.empty_to_none(body.get("recordId")))
        hod_email = str(m.empty_to_none(body.get("hodEmail")) or "").strip()
        if not record_id:
            raise HTTPException(status_code=400, detail="recordId is required")
        if not hod_email:
            raise HTTPException(status_code=400, detail="hodEmail is required")
        if "@" not in hod_email or "." not in hod_email.split("@", 1)[-1]:
            raise HTTPException(status_code=400, detail="Invalid hodEmail format")

        local_record, sql_record = get_customer_records(record_id)
        if not local_record and not sql_record:
            raise HTTPException(status_code=404, detail="Customer record not found")

        user_email = str(user.get("email") or "").strip().lower()
        if not user_email:
            raise HTTPException(status_code=403, detail="Access denied for this record")

        assigned_to_user = (
            m._record_is_assigned_to_email(local_record, user_email)
            or m._record_is_assigned_to_email(sql_record, user_email)
        )
        if not assigned_to_user:
            for row in customer_db.get_customer_records_by_assignee_email(user_email):
                row_record = m._normalize_record_id(row.get("record"))
                if row_record and row_record.lower() == record_id.lower():
                    assigned_to_user = True
                    break
        if not assigned_to_user:
            raise HTTPException(status_code=403, detail="Access denied for this record")

        current_status = str(
            m._first_existing_record_value(
                sql_record,
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
        if current_status == "pending for sap code creation":
            raise HTTPException(status_code=409, detail="HOD email cannot be updated after SAP stage")

        updates = {
            "hodEmail": hod_email,
            "HodEmail": hod_email,
            "HODEmail": hod_email,
        }
        db_updated = customer_db.update_customer_record_by_record_id(record_id, updates)
        if sql_record and not db_updated:
            raise HTTPException(status_code=500, detail="Customer record found in database but update failed")

        local_updated = m.update_local_vendor_record_by_record_id(record_id, updates)
        if not local_updated and not db_updated:
            raise HTTPException(status_code=400, detail="Unable to update HOD email")

        return {
            "message": "HOD email updated successfully",
            "recordId": record_id,
            "hodEmail": hod_email,
        }

    @app.post("/api/customer/hod-bulk-decision")
    def submit_customer_hod_bulk_decision(
        body: dict[str, Any],
        request: Request,
        user: dict[str, Any] = Depends(m.get_current_user),
    ) -> dict[str, Any]:
        if not m.user_has_any_role(user, "hod"):
            raise HTTPException(status_code=403, detail="Only HOD can use this action")

        raw_record_ids = body.get("recordIds")
        decision = str(m.empty_to_none(body.get("decision")) or "").strip().lower()
        remarks = str(m.empty_to_none(body.get("remarks")) or "").strip()
        reject_to = str(m.empty_to_none(body.get("rejectTo")) or "").strip()
        reject_target = normalize_customer_reject_target(reject_to)
        user_email = str(user.get("email") or "").strip().lower()
        user_role = str(user.get("role") or "HOD").strip() or "HOD"

        if not isinstance(raw_record_ids, list) or not raw_record_ids:
            raise HTTPException(status_code=400, detail="recordIds must be a non-empty list")
        if decision not in {"approve", "reject"}:
            raise HTTPException(status_code=400, detail="decision must be Approve or Reject")
        if decision == "reject" and reject_target not in {"customer", "buyer"}:
            raise HTTPException(status_code=400, detail="rejectTo must be Customer or Buyer")
        if not remarks:
            raise HTTPException(status_code=400, detail="remarks are required")
        if not user_email:
            raise HTTPException(status_code=400, detail="User email not found")

        normalized_record_ids: list[str] = []
        seen_record_ids: set[str] = set()
        for raw_id in raw_record_ids:
            record_id = m._normalize_record_id(raw_id)
            if not record_id:
                continue
            if record_id in seen_record_ids:
                continue
            seen_record_ids.add(record_id)
            normalized_record_ids.append(record_id)
        if not normalized_record_ids:
            raise HTTPException(status_code=400, detail="No valid recordIds provided")

        pending_status_keys = {"pending for hod approval", "pendign for hod approval"}
        target_status = STATUS_PENDING_SAP_CODE_CREATION if decision == "approve" else "Rejected by HOD"
        decision_label = "Approve" if decision == "approve" else "Reject"
        updated_record_ids: list[str] = []
        failed_records: list[dict[str, str]] = []
        rejection_email_sent_count = 0
        rejection_email_errors: list[str] = []

        for record_id in normalized_record_ids:
            local_record, sql_record = get_customer_records(record_id)
            if not local_record and not sql_record:
                failed_records.append({"recordId": record_id, "reason": "Customer record not found"})
                continue

            assigned_hod_email = m._extract_record_hod_email(local_record) or m._extract_record_hod_email(sql_record)
            if not assigned_hod_email or assigned_hod_email != user_email:
                failed_records.append({"recordId": record_id, "reason": "Record is not assigned to current HOD"})
                continue

            status_value = str(
                m._first_existing_record_value(
                    sql_record,
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
                "reviewRemarks": remarks,
                "ReviewRemarks": remarks,
                "reviewedByRole": user_role,
                "ReviewedByRole": user_role,
                "remarks": remarks,
                "Remarks": remarks,
            }

            db_updated = customer_db.update_customer_record_by_record_id(record_id, updates)
            local_updated = m.update_local_vendor_record_by_record_id(record_id, updates)
            if not db_updated and not local_updated:
                failed_records.append({"recordId": record_id, "reason": "Unable to update record"})
                continue

            updated_record_ids.append(record_id)

            if decision == "reject":
                update_customer_form_url = build_customer_public_url("/web/update-customer", request=request)
                buyer_email = (
                    m._extract_record_buyer_email(local_record)
                    or m._extract_record_buyer_email(sql_record)
                )
                buyer_name = (
                    m._extract_record_buyer_name(local_record)
                    or m._extract_record_buyer_name(sql_record)
                )
                customer_email = (
                    m._extract_record_vendor_email(local_record)
                    or m._extract_record_vendor_email(sql_record)
                )
                customer_name = get_customer_display_name(local_record, sql_record)
                try:
                    if reject_target == "customer":
                        if not customer_email:
                            raise RuntimeError("Customer email not found")
                        m.send_vendor_rejection_email(
                            recipient_name=customer_name,
                            recipient_email=customer_email,
                            record_id=record_id,
                            remarks=remarks,
                            update_form_url=update_customer_form_url,
                            rejected_by="HOD",
                            cc_email=buyer_email or None,
                            entity_label="Customer",
                        )
                    else:
                        if not buyer_email:
                            raise RuntimeError("Buyer email not found")
                        m.send_vendor_rejection_email(
                            recipient_name=buyer_name,
                            recipient_email=buyer_email,
                            record_id=record_id,
                            remarks=remarks,
                            update_form_url=update_customer_form_url,
                            rejected_by="HOD",
                            entity_label="Customer",
                        )
                    rejection_email_sent_count += 1
                except Exception as exc:
                    rejection_email_errors.append(f"{record_id}: {exc}")

        if not updated_record_ids:
            raise HTTPException(status_code=400, detail="Unable to update selected records for HOD decision")

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
            "sapSyncStarted": False,
            "sapSyncQueuedCount": 0,
        }
