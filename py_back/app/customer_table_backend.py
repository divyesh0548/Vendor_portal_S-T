from __future__ import annotations

from datetime import datetime, timedelta, timezone
from typing import Any

from psycopg2 import IntegrityError

from . import db as shared_db

CUSTOMER_TABLE_NAMES = [
    "CustomerMasterData",
    "Customer Master Data",
    "CustomerFormUnified",
]

CUSTOMER_RECORD_COLS = ["recordId", "RecordID", "record", "Record", "Id", "ID"]
CUSTOMER_STATUS_COLS = [
    "approverStatus",
    "ApproverStatus",
    "ApprovStatus",
    "status",
    "Status",
    "inviteStatus",
    "InviteStatus",
]
CUSTOMER_CREATED_AT_COLS = ["created_at", "updated_at", "createdAt", "CreatedAt", "updatedAt", "UpdatedAt"]
CUSTOMER_COMPANY_COLS = [
    "companyName",
    "CompanyName",
    "CompanyDealing",
    "CompanyDealingWith",
    "company",
    "CustomerGroup",
]
CUSTOMER_CATEGORY_COLS = [
    "FirmType",
    "firmType",
    "CustomerGroup",
    "customerGroup",
    "vendorCategory",
    "VendorCategory",
]
CUSTOMER_NAME_COLS = [
    "CustomerName",
    "customerName",
    "vendorName",
    "VendorName",
    "TopCustomer",
    "name",
]
CUSTOMER_TYPE_COLS = [
    "CustomerType",
    "vendorType",
    "VendorType",
    "TypeofEntity",
    "IndustryType",
    "type",
]
CUSTOMER_ASSIGNEE_EMAIL_COLS = [
    "Rubamin Contact Person Email",
    "rubaminContactPersonEmail",
    "Rubamin Contact Person",
    "rubaminContactPerson",
    "invitedByEmail",
    "InvitedByEmail",
    "Invited By Email",
    "contactPersonEmail",
    "ContactPersonEmail",
    "ContactEmailID",
    "contactEmailId",
    "customerEmail",
    "CustomerEmail",
    "Email",
]
CUSTOMER_HOD_EMAIL_COLS = [
    "rubaminApproverHod",
    "RubaminApproverHod",
    "RubaminApprovalHod",
    "hodEmail",
    "HodEmail",
    "HODEmail",
]

CUSTOMER_CANONICAL_COLUMN_ALIASES: dict[str, list[str]] = {
    "approverStatus": [
        "ApproverStatus",
        "ApprovStatus",
        "status",
        "Status",
        "inviteStatus",
        "InviteStatus",
    ],
    "CustomerName": [
        "customerName",
        "vendorName",
        "VendorName",
        "Vendorname",
    ],
    "CustomerGroup": [
        "customerGroup",
        "vendorCategory",
        "VendorCategory",
        "Vendor_Category",
        "companyName",
        "CompanyName",
        "company",
    ],
    "CustomerType": [
        "vendorType",
        "VendorType",
    ],
    "customerEmail": [
        "CustomerEmail",
        "vendorEmail",
        "VendorEmail",
    ],
    "created_at": [
        "createdAt",
        "CreatedAt",
    ],
    "updated_at": [
        "updatedAt",
        "UpdatedAt",
    ],
    "rubaminContactPersonEmail": [
        "Rubamin Contact Person Email",
        "invitedByEmail",
        "InvitedByEmail",
        "Invited By Email",
        "contactPersonEmail",
        "ContactPersonEmail",
        "ContactEmailID",
        "contactEmailId",
    ],
    "rubaminContactPerson": [
        "Rubamin Contact Person",
    ],
    "rubaminContactPersonDepartment": [
        "Rubamin Contact Person Department",
    ],
    "hodEmail": [
        "rubaminApproverHod",
        "RubaminApproverHod",
        "RubaminApprovalHod",
        "HodEmail",
        "HODEmail",
    ],
    "GST Valid": [
        "gstValid",
        "GSTValid",
    ],
    "TaxpayerTypeGST": [
        "taxpayerTypeGst",
        "TaxpayerType",
        "taxpayerType",
        "Taxpayer Type",
    ],
    "reviewRemarks": [
        "ReviewRemarks",
    ],
    "reviewedByRole": [
        "ReviewedByRole",
    ],
    "remarks": [
        "Remarks",
    ],
    "withholdingTax": [
        "WithholdingTax",
        "Withholding Tax",
    ],
    "tcsRate": [
        "TCSRate",
        "TCS Rate",
    ],
    "turnoverLimit": [
        "TurnoverLimit",
        "Turnover Limit",
    ],
    "gstValidate": [
        "GST Validate",
    ],
    "cinValidate": [
        "CIN Validate",
    ],
    "panValidate": [
        "PAN Validate",
    ],
}

CUSTOMER_COLUMN_CANONICAL_MAP: dict[str, str] = {}
for _canonical_col, _alias_cols in CUSTOMER_CANONICAL_COLUMN_ALIASES.items():
    CUSTOMER_COLUMN_CANONICAL_MAP[_canonical_col.lower()] = _canonical_col
    for _alias_col in _alias_cols:
        CUSTOMER_COLUMN_CANONICAL_MAP[str(_alias_col).strip().lower()] = _canonical_col

DUPLICATE_CUSTOMER_COLUMN_PAIRS: list[tuple[str, str]] = [
    ("approverStatus", "ApproverStatus"),
    ("approverStatus", "ApprovStatus"),
    ("approverStatus", "status"),
    ("approverStatus", "Status"),
    ("CustomerName", "customerName"),
    ("CustomerName", "vendorName"),
    ("CustomerName", "VendorName"),
    ("CustomerName", "Vendorname"),
    ("CustomerGroup", "customerGroup"),
    ("CustomerGroup", "vendorCategory"),
    ("CustomerGroup", "VendorCategory"),
    ("CustomerGroup", "Vendor_Category"),
    ("CustomerType", "vendorType"),
    ("CustomerType", "VendorType"),
    ("customerEmail", "CustomerEmail"),
    ("customerEmail", "vendorEmail"),
    ("customerEmail", "VendorEmail"),
    ("TaxpayerTypeGST", "taxpayerTypeGst"),
    ("TaxpayerTypeGST", "TaxpayerType"),
    ("rubaminContactPersonEmail", "Rubamin Contact Person Email"),
    ("rubaminContactPerson", "Rubamin Contact Person"),
    ("rubaminContactPersonDepartment", "Rubamin Contact Person Department"),
    ("created_at", "createdAt"),
    ("created_at", "CreatedAt"),
    ("updated_at", "updatedAt"),
    ("updated_at", "UpdatedAt"),
    ("reviewRemarks", "ReviewRemarks"),
    ("reviewedByRole", "ReviewedByRole"),
    ("remarks", "Remarks"),
    ("withholdingTax", "WithholdingTax"),
    ("withholdingTax", "Withholding Tax"),
    ("tcsRate", "TCSRate"),
    ("tcsRate", "TCS Rate"),
    ("turnoverLimit", "TurnoverLimit"),
    ("turnoverLimit", "Turnover Limit"),
    ("gstValidate", "GST Validate"),
    ("cinValidate", "CIN Validate"),
    ("panValidate", "PAN Validate"),
]

# Kept for backward compatibility; no GST columns are obsolete now.
OBSOLETE_CUSTOMER_GST_COLUMNS: set[str] = set()

REMOVED_UNUSED_CUSTOMER_COLUMNS = {
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
    "contactpersonemail",
    "contactemailid",
    "rubaminapproverhod",
    "rubaminapprovalhod",
    "buyeremail",
}

CUSTOMER_WORKFLOW_COLUMN_TYPES: dict[str, str] = {
    "approverStatus": "VARCHAR(200) NULL",
    "created_at": "VARCHAR(40) NULL",
    "updated_at": "VARCHAR(40) NULL",
    "customerEmail": "VARCHAR(320) NULL",
    "rubaminContactPersonEmail": "VARCHAR(320) NULL",
    "rubaminContactPerson": "VARCHAR(320) NULL",
    "rubaminContactPersonDepartment": "VARCHAR(320) NULL",
    "hodEmail": "VARCHAR(320) NULL",
    "HODEmail": "VARCHAR(320) NULL",
    "CustomerName": "VARCHAR(500) NULL",
    "CustomerGroup": "VARCHAR(500) NULL",
    "CustomerType": "VARCHAR(500) NULL",
    "reviewRemarks": "TEXT NULL",
    "reviewedByRole": "VARCHAR(100) NULL",
    "remarks": "TEXT NULL",
    "withholdingTax": "TEXT NULL",
    "tcsRate": "VARCHAR(200) NULL",
    "turnoverLimit": "VARCHAR(200) NULL",
    "gstValidate": "VARCHAR(5) NULL",
    "GST Valid": "VARCHAR(30) NULL",
    "TaxpayerTypeGST": "VARCHAR(200) NULL",
    "cinValidate": "VARCHAR(5) NULL",
    "panValidate": "VARCHAR(5) NULL",
}


def _now_str() -> str:
    ist = timezone(timedelta(hours=5, minutes=30))
    return datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S")


def _ensure_connection() -> Any | None:
    if not shared_db.conn:
        shared_db.initialize_database()
    return shared_db.conn


def _find_customer_table(cursor: Any) -> tuple[str, str] | None:
    return shared_db._find_table(cursor, CUSTOMER_TABLE_NAMES)


def _resolve_customer_table(cursor: Any) -> tuple[str, str, str] | None:
    located = _find_customer_table(cursor)
    if not located:
        return None
    schema, table = located
    return schema, table, shared_db._qt(schema, table)


def _safe_column_name(column_name: str) -> bool:
    value = str(column_name or "").strip()
    if not value:
        return False
    if '"' in value:
        return False
    return True


def _is_obsolete_customer_column(column_name: str) -> bool:
    key = str(column_name or "").strip().lower()
    if not key:
        return False
    if key in OBSOLETE_CUSTOMER_GST_COLUMNS:
        return True
    return key in REMOVED_UNUSED_CUSTOMER_COLUMNS


def _pick_value(row: dict[str, Any], columns: list[str], default: str = "") -> str:
    for col in columns:
        value = row.get(col)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return default


def _resolve_column_name(row: dict[str, Any], candidate: str) -> str | None:
    lowered_candidate = str(candidate or "").strip().lower()
    if not lowered_candidate:
        return None
    for key in row.keys():
        if str(key).strip().lower() == lowered_candidate:
            return str(key)
    return None


def _is_empty_text_value(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def _set_missing_alias_value(row: dict[str, Any], key: str, value: Any) -> None:
    existing = row.get(key)
    if key in row and not _is_empty_text_value(existing):
        return
    row[key] = value


def _hydrate_customer_alias_fields(row: dict[str, Any]) -> dict[str, Any]:
    status_value = _pick_value(row, ["approverStatus", "ApproverStatus", "ApprovStatus", "status", "Status"], "")
    if status_value:
        _set_missing_alias_value(row, "approverStatus", status_value)
        _set_missing_alias_value(row, "status", status_value)
        _set_missing_alias_value(row, "ApproverStatus", status_value)
        _set_missing_alias_value(row, "ApprovStatus", status_value)

    name_value = _pick_value(row, ["CustomerName", "customerName", "vendorName", "VendorName", "Vendorname"], "")
    if name_value:
        _set_missing_alias_value(row, "CustomerName", name_value)
        _set_missing_alias_value(row, "vendorName", name_value)
        _set_missing_alias_value(row, "VendorName", name_value)

    category_value = _pick_value(
        row,
        ["CustomerGroup", "customerGroup", "vendorCategory", "VendorCategory", "Vendor_Category", "FirmType", "firmType"],
        "",
    )
    if category_value:
        _set_missing_alias_value(row, "CustomerGroup", category_value)
        _set_missing_alias_value(row, "vendorCategory", category_value)
        _set_missing_alias_value(row, "VendorCategory", category_value)

    type_value = _pick_value(row, ["CustomerType", "vendorType", "VendorType"], "")
    if type_value:
        _set_missing_alias_value(row, "CustomerType", type_value)
        _set_missing_alias_value(row, "vendorType", type_value)
        _set_missing_alias_value(row, "VendorType", type_value)

    email_value = _pick_value(row, ["customerEmail", "CustomerEmail", "vendorEmail", "VendorEmail"], "")
    if email_value:
        _set_missing_alias_value(row, "customerEmail", email_value)
        _set_missing_alias_value(row, "vendorEmail", email_value)
        _set_missing_alias_value(row, "VendorEmail", email_value)

    taxpayer_type_value = _pick_value(
        row,
        ["TaxpayerTypeGST", "taxpayerTypeGst", "TaxpayerType", "taxpayerType", "Taxpayer Type"],
        "",
    )
    if taxpayer_type_value:
        _set_missing_alias_value(row, "TaxpayerTypeGST", taxpayer_type_value)
        _set_missing_alias_value(row, "taxpayerTypeGst", taxpayer_type_value)
        _set_missing_alias_value(row, "TaxpayerType", taxpayer_type_value)

    created_at_value = _pick_value(row, ["createdAt", "CreatedAt", "created_at"], "")
    if created_at_value:
        _set_missing_alias_value(row, "createdAt", created_at_value)
        _set_missing_alias_value(row, "CreatedAt", created_at_value)
        _set_missing_alias_value(row, "created_at", created_at_value)

    updated_at_value = _pick_value(row, ["updatedAt", "UpdatedAt", "updated_at"], "")
    if updated_at_value:
        _set_missing_alias_value(row, "updatedAt", updated_at_value)
        _set_missing_alias_value(row, "UpdatedAt", updated_at_value)
        _set_missing_alias_value(row, "updated_at", updated_at_value)

    rubamin_email_value = _pick_value(
        row,
        [
            "rubaminContactPersonEmail",
            "Rubamin Contact Person Email",
            "invitedByEmail",
            "InvitedByEmail",
            "Invited By Email",
            "contactPersonEmail",
            "ContactPersonEmail",
            "ContactEmailID",
            "contactEmailId",
        ],
        "",
    )
    if rubamin_email_value:
        _set_missing_alias_value(row, "rubaminContactPersonEmail", rubamin_email_value)
        _set_missing_alias_value(row, "Rubamin Contact Person Email", rubamin_email_value)

    rubamin_contact_name_value = _pick_value(
        row,
        [
            "rubaminContactPerson",
            "Rubamin Contact Person",
        ],
        "",
    )
    if rubamin_contact_name_value:
        _set_missing_alias_value(row, "rubaminContactPerson", rubamin_contact_name_value)
        _set_missing_alias_value(row, "Rubamin Contact Person", rubamin_contact_name_value)

    rubamin_contact_department_value = _pick_value(
        row,
        [
            "rubaminContactPersonDepartment",
            "Rubamin Contact Person Department",
        ],
        "",
    )
    if rubamin_contact_department_value:
        _set_missing_alias_value(row, "rubaminContactPersonDepartment", rubamin_contact_department_value)
        _set_missing_alias_value(row, "Rubamin Contact Person Department", rubamin_contact_department_value)

    hod_email_value = _pick_value(
        row,
        [
            "hodEmail",
            "HodEmail",
            "HODEmail",
            "rubaminApproverHod",
            "RubaminApproverHod",
            "RubaminApprovalHod",
        ],
        "",
    )
    if hod_email_value:
        _set_missing_alias_value(row, "hodEmail", hod_email_value)
        _set_missing_alias_value(row, "HodEmail", hod_email_value)
        _set_missing_alias_value(row, "HODEmail", hod_email_value)

    return row


def _canonicalize_customer_column_name(column_name: str) -> str:
    raw = str(column_name or "").strip()
    if not raw:
        return ""
    return CUSTOMER_COLUMN_CANONICAL_MAP.get(raw.lower(), raw)


def _normalize_customer_payload_keys(payload: dict[str, Any]) -> dict[str, Any]:
    normalized: dict[str, Any] = {}
    deferred_alias_values: list[tuple[str, Any]] = []

    for raw_key, raw_value in payload.items():
        key = str(raw_key or "").strip()
        if not key:
            continue

        canonical_key = _canonicalize_customer_column_name(key)
        if not canonical_key:
            continue
        if _is_obsolete_customer_column(canonical_key):
            continue

        if canonical_key.lower() == key.lower():
            if canonical_key not in normalized or (
                _is_empty_text_value(normalized.get(canonical_key)) and not _is_empty_text_value(raw_value)
            ):
                normalized[canonical_key] = raw_value
        else:
            deferred_alias_values.append((canonical_key, raw_value))

    for canonical_key, raw_value in deferred_alias_values:
        if canonical_key not in normalized or (
            _is_empty_text_value(normalized.get(canonical_key)) and not _is_empty_text_value(raw_value)
        ):
            normalized[canonical_key] = raw_value

    return normalized


def _resolve_payload_column(col_map: dict[str, str], payload_key: str) -> str | None:
    raw_key = str(payload_key or "").strip()
    if not raw_key:
        return None

    canonical_key = _canonicalize_customer_column_name(raw_key)
    candidates = [canonical_key]
    for alias in CUSTOMER_CANONICAL_COLUMN_ALIASES.get(canonical_key, []):
        candidates.append(alias)

    for candidate in candidates:
        actual = col_map.get(str(candidate).strip().lower())
        if actual:
            return actual
    return None


def _coerce_binary_value(value: Any) -> bytes | None:
    if value is None:
        return None
    if isinstance(value, bytes):
        return value
    if isinstance(value, bytearray):
        return bytes(value)
    if isinstance(value, memoryview):
        return value.tobytes()
    return None


def _summary_row(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "record": _pick_value(row, CUSTOMER_RECORD_COLS),
        "company": _pick_value(row, CUSTOMER_COMPANY_COLS),
        "category": _pick_value(row, CUSTOMER_CATEGORY_COLS),
        "approverStatus": _pick_value(row, CUSTOMER_STATUS_COLS),
        "name": _pick_value(row, CUSTOMER_NAME_COLS),
        "type": _pick_value(row, CUSTOMER_TYPE_COLS),
        "createdAt": _pick_value(row, CUSTOMER_CREATED_AT_COLS),
    }


def ensure_customer_columns(column_names: list[str], default_type: str = "TEXT NULL") -> list[str]:
    connection = _ensure_connection()
    if not connection:
        raise RuntimeError(shared_db.db_status.get("error") or "Database connection failed")

    try:
        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            raise RuntimeError("Customer master table not found")

        schema, table, full_table = table_info
        existing_columns = shared_db._list_columns(cursor, schema, table)
        existing_lower = {column.lower() for column in existing_columns}
        added_columns: list[str] = []

        for raw_name in column_names:
            column_name = str(raw_name or "").strip()
            if not _safe_column_name(column_name):
                continue
            if column_name.lower() in existing_lower:
                continue

            cursor.execute(f"ALTER TABLE {full_table} ADD COLUMN {shared_db._qi(column_name)} {default_type}")
            existing_lower.add(column_name.lower())
            added_columns.append(column_name)

        if added_columns:
            connection.commit()
        return added_columns
    except Exception as exc:
        try:
            connection.rollback()
        except Exception:
            pass
        raise RuntimeError(f"Unable to ensure customer columns: {exc}") from exc


def ensure_customer_workflow_columns() -> list[str]:
    connection = _ensure_connection()
    if not connection:
        raise RuntimeError(shared_db.db_status.get("error") or "Database connection failed")

    try:
        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            raise RuntimeError("Customer master table not found")

        schema, table, full_table = table_info
        existing_columns = shared_db._list_columns(cursor, schema, table)
        existing_lower = {column.lower() for column in existing_columns}
        added_columns: list[str] = []

        for column_name, column_type in CUSTOMER_WORKFLOW_COLUMN_TYPES.items():
            if not _safe_column_name(column_name):
                continue
            if column_name.lower() in existing_lower:
                continue
            cursor.execute(f"ALTER TABLE {full_table} ADD COLUMN {shared_db._qi(column_name)} {column_type}")
            existing_lower.add(column_name.lower())
            added_columns.append(column_name)

        if added_columns:
            connection.commit()
        return added_columns
    except Exception as exc:
        try:
            connection.rollback()
        except Exception:
            pass
        raise RuntimeError(f"Unable to ensure customer workflow columns: {exc}") from exc


def remove_obsolete_customer_gst_columns() -> dict[str, Any]:
    connection = _ensure_connection()
    if not connection:
        return {
            "table": None,
            "removedColumns": [],
            "missingColumns": [],
            "error": shared_db.db_status.get("error") or "Database connection failed",
        }

    try:
        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            return {
                "table": None,
                "removedColumns": [],
                "missingColumns": [],
                "error": "Customer master table not found",
            }

        schema, table, full_table = table_info
        existing_columns = shared_db._list_columns(cursor, schema, table)
        existing_map = {str(column).strip().lower(): str(column) for column in existing_columns}

        removable_actual_columns: list[str] = []
        missing_columns: list[str] = []
        for candidate in sorted(OBSOLETE_CUSTOMER_GST_COLUMNS):
            actual_column = existing_map.get(candidate)
            if actual_column:
                removable_actual_columns.append(actual_column)
            else:
                missing_columns.append(candidate)

        for actual_column in removable_actual_columns:
            cursor.execute(f"ALTER TABLE {full_table} DROP COLUMN {shared_db._qi(actual_column)}")

        if removable_actual_columns:
            connection.commit()

        return {
            "table": full_table,
            "removedColumns": removable_actual_columns,
            "missingColumns": missing_columns,
        }
    except Exception as exc:
        try:
            connection.rollback()
        except Exception:
            pass
        raise RuntimeError(f"Unable to remove obsolete customer GST columns: {exc}") from exc


def _cleanup_duplicate_customer_columns(
    cursor: Any,
    schema: str,
    table: str,
) -> list[str]:
    dropped_columns: list[str] = []
    columns = shared_db._list_columns(cursor, schema, table)

    for canonical_name, duplicate_name in DUPLICATE_CUSTOMER_COLUMN_PAIRS:
        canonical_col = shared_db._resolve_column_name(columns, canonical_name)
        duplicate_col = shared_db._resolve_column_name(columns, duplicate_name)
        if not canonical_col or not duplicate_col:
            continue
        if canonical_col == duplicate_col:
            continue

        try:
            cursor.execute(
                f"""
                UPDATE {shared_db._qt(schema, table)}
                SET {shared_db._qi(canonical_col)} = {shared_db._qi(duplicate_col)}
                WHERE {shared_db._qi(duplicate_col)} IS NOT NULL
                  AND (
                    {shared_db._qi(canonical_col)} IS NULL
                    OR TRIM(CAST({shared_db._qi(canonical_col)} AS TEXT)) = ''
                  )
                """
            )
            cursor.execute(f"ALTER TABLE {shared_db._qt(schema, table)} DROP COLUMN {shared_db._qi(duplicate_col)}")
            dropped_columns.append(duplicate_col)
            columns.discard(duplicate_col)
        except Exception:
            # Keep cleanup safe even if drop fails due constraints.
            continue

    return dropped_columns


def _drop_removed_customer_columns(
    cursor: Any,
    schema: str,
    table: str,
) -> list[str]:
    dropped_columns: list[str] = []
    columns = shared_db._list_columns(cursor, schema, table)
    lowered_map = {str(column).strip().lower(): str(column) for column in columns}

    for candidate in sorted(REMOVED_UNUSED_CUSTOMER_COLUMNS):
        actual_column = lowered_map.get(candidate)
        if not actual_column:
            continue
        try:
            cursor.execute(f"ALTER TABLE {shared_db._qt(schema, table)} DROP COLUMN {shared_db._qi(actual_column)}")
            dropped_columns.append(actual_column)
            lowered_map.pop(candidate, None)
        except Exception:
            # Keep cleanup safe even if a column cannot be dropped due constraints.
            continue

    return dropped_columns


def cleanup_customer_duplicate_columns() -> dict[str, Any]:
    connection = _ensure_connection()
    if not connection:
        return {
            "success": False,
            "droppedColumns": [],
            "error": shared_db.db_status.get("error") or "Database connection failed",
        }

    try:
        # Ensure canonical columns exist before migrating and dropping aliases.
        ensure_customer_workflow_columns()

        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            return {
                "success": False,
                "droppedColumns": [],
                "error": "Customer master table not found",
            }

        schema, table, _ = table_info
        dropped_duplicates = _cleanup_duplicate_customer_columns(cursor, schema, table)
        dropped_removed = _drop_removed_customer_columns(cursor, schema, table)
        dropped = dropped_duplicates + dropped_removed
        connection.commit()
        return {
            "success": True,
            "schema": schema,
            "table": table,
            "droppedColumns": dropped,
            "droppedDuplicateColumns": dropped_duplicates,
            "droppedRemovedColumns": dropped_removed,
        }
    except Exception as exc:
        try:
            connection.rollback()
        except Exception:
            pass
        return {
            "success": False,
            "droppedColumns": [],
            "error": str(exc),
        }


def _ensure_payload_columns(
    cursor: Any,
    schema: str,
    table: str,
    payload: dict[str, Any],
) -> tuple[set[str], dict[str, str], set[str], dict[str, str]]:
    payload = _normalize_customer_payload_keys(payload)
    columns = shared_db._list_columns(cursor, schema, table)
    col_map = {column.lower(): column for column in columns}
    missing_columns = [
        key
        for key in payload.keys()
        if (
            _safe_column_name(key)
            and not _is_obsolete_customer_column(str(key))
            and str(key).strip().lower() not in col_map
        )
    ]
    if missing_columns:
        ensure_customer_columns(missing_columns)
        columns = shared_db._list_columns(cursor, schema, table)
        col_map = {column.lower(): column for column in columns}

    col_types = shared_db._list_column_types(cursor, schema, table)
    blocked_columns = shared_db._list_non_insertable_columns(cursor, schema, table)
    return columns, col_types, blocked_columns, col_map


def insert_customer_master_data(customer_data: dict[str, Any]) -> str:
    connection = _ensure_connection()
    if not connection:
        raise RuntimeError(shared_db.db_status.get("error") or "Database connection failed")

    ensure_customer_workflow_columns()

    cursor = connection.cursor()
    table_info = _resolve_customer_table(cursor)
    if not table_info:
        raise RuntimeError("Customer master table not found")

    schema, table, full_table = table_info

    normalized_customer_data = _normalize_customer_payload_keys(customer_data)
    columns, col_types, blocked_columns, col_map = _ensure_payload_columns(
        cursor, schema, table, normalized_customer_data
    )
    record_col = shared_db._pick_col(columns, CUSTOMER_RECORD_COLS)
    if record_col and record_col.lower() in blocked_columns:
        record_col = None
    record_col_type = col_types.get(record_col.lower(), "") if record_col else ""

    payload: dict[str, Any] = {}
    for key, value in normalized_customer_data.items():
        if value is None:
            continue
        if isinstance(value, (list, dict, tuple, set)):
            value = str(value)

        key_lower = str(key or "").strip().lower()
        if not key_lower:
            continue
        if _is_obsolete_customer_column(key_lower):
            continue
        actual_col = _resolve_payload_column(col_map, key)
        if not actual_col:
            continue
        if actual_col.lower() in blocked_columns:
            continue

        sql_type = col_types.get(actual_col.lower(), "")
        payload[actual_col] = shared_db._coerce_sql_value(value, sql_type)

    now_str = _now_str()
    for timestamp_key in ("created_at", "updated_at"):
        actual_col = col_map.get(timestamp_key.lower())
        if actual_col and actual_col.lower() not in blocked_columns and actual_col not in payload:
            sql_type = col_types.get(actual_col.lower(), "")
            payload[actual_col] = shared_db._coerce_sql_value(now_str, sql_type) or now_str

    requested_record_id = ""
    if record_col and record_col in payload:
        requested_record_id = str(payload.get(record_col) or "").strip()

    generated_record_id = ""
    last_error: Exception | None = None
    for attempt in range(6):
        local_payload = dict(payload)
        if record_col:
            if requested_record_id and attempt == 0:
                generated_record_id = requested_record_id
            elif record_col_type in {"uniqueidentifier", "uuid"}:
                generated_record_id = str(shared_db.uuid.uuid4())
            else:
                generated_record_id = shared_db._generate_record_id(24)
            local_payload[record_col] = generated_record_id

        if not local_payload:
            raise RuntimeError("No matching columns found to insert")

        cols_sql = ", ".join(f"{shared_db._qi(column)}" for column in local_payload.keys())
        params_sql = ", ".join("%s" for _ in local_payload.values())
        values = list(local_payload.values())

        try:
            cursor.execute(f"INSERT INTO {full_table} ({cols_sql}) VALUES ({params_sql})", values)
            connection.commit()
            return generated_record_id
        except IntegrityError as exc:
            last_error = exc
            message = str(exc).lower()
            if "duplicate" in message or "unique" in message or "2627" in message or "2601" in message:
                continue
            raise RuntimeError(f"SQL integrity error: {exc}") from exc
        except Exception as exc:
            raise RuntimeError(f"SQL insert error: {exc}") from exc

    if last_error:
        raise RuntimeError(f"Unable to insert customer data after retries: {last_error}") from last_error
    raise RuntimeError("Unable to insert customer data after retries")


def update_customer_record_by_record_id(record_id: str, updates: dict[str, Any]) -> bool:
    record_value = str(record_id or "").strip()
    if not record_value or not isinstance(updates, dict) or not updates:
        return False

    connection = _ensure_connection()
    if not connection:
        return False

    ensure_customer_workflow_columns()

    try:
        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            return False

        schema, table, full_table = table_info
        normalized_updates = _normalize_customer_payload_keys(updates)
        columns, col_types, blocked_columns, col_map = _ensure_payload_columns(
            cursor, schema, table, normalized_updates
        )
        record_col = shared_db._pick_col(columns, CUSTOMER_RECORD_COLS)
        if not record_col:
            return False

        payload: dict[str, Any] = {}
        for key, value in normalized_updates.items():
            key_lower = str(key or "").strip().lower()
            if not key_lower:
                continue
            if _is_obsolete_customer_column(key_lower):
                continue

            actual_col = _resolve_payload_column(col_map, key)
            if not actual_col:
                continue
            if actual_col.lower() in blocked_columns:
                continue
            if actual_col.lower() == record_col.lower():
                continue

            if isinstance(value, (list, dict, tuple, set)):
                value = str(value)
            if isinstance(value, str):
                value = value.strip()

            sql_type = col_types.get(actual_col.lower(), "")
            payload[actual_col] = shared_db._coerce_sql_value(value, sql_type)

        for timestamp_key in ("updated_at",):
            actual_col = col_map.get(timestamp_key.lower())
            if not actual_col:
                continue
            if actual_col.lower() in blocked_columns or actual_col.lower() == record_col.lower():
                continue
            payload[actual_col] = _now_str()
            break

        if not payload:
            return False

        set_parts = [f"{shared_db._qi(column)} = %s" for column in payload.keys()]
        params = list(payload.values())
        params.append(record_value)

        query = (
            f"UPDATE {full_table} "
            f"SET {', '.join(set_parts)} "
            f"WHERE LOWER(TRIM(CAST({shared_db._qi(record_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s))"
        )
        cursor.execute(query, params)
        affected_rows = cursor.rowcount
        connection.commit()
        return affected_rows > 0
    except Exception:
        try:
            connection.rollback()
        except Exception:
            pass
        return False


def update_customer_approver_status_by_record_id(record_id: str, approver_status: str) -> bool:
    updates = {"approverStatus": approver_status}
    return update_customer_record_by_record_id(record_id, updates)


def get_customer_record_by_record_id(record_id: str) -> dict[str, Any] | None:
    record_value = str(record_id or "").strip()
    if not record_value:
        return None

    connection = _ensure_connection()
    if not connection:
        return None

    try:
        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            return None

        schema, table, full_table = table_info
        columns = shared_db._list_columns(cursor, schema, table)
        record_col = shared_db._pick_col(columns, CUSTOMER_RECORD_COLS)
        if not record_col:
            return None

        query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({shared_db._qi(record_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s)) LIMIT 1"
        )
        cursor.execute(query, (record_value,))
        row = cursor.fetchone()
        if not row:
            return None

        row_dict: dict[str, Any] = {}
        for idx, desc in enumerate(cursor.description):
            key = str(desc[0])
            row_dict[key] = shared_db._to_json_safe_value(row[idx])
        return _hydrate_customer_alias_fields(row_dict)
    except Exception:
        return None


def get_all_customer_records() -> list[dict[str, Any]]:
    connection = _ensure_connection()
    if not connection:
        return []

    try:
        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            return []

        _, _, full_table = table_info
        cursor.execute(f"SELECT * FROM {full_table}")
        rows = cursor.fetchall()
        description = cursor.description

        records: list[dict[str, Any]] = []
        for row in rows:
            row_dict: dict[str, Any] = {}
            for idx, desc in enumerate(description):
                row_dict[str(desc[0])] = shared_db._to_json_safe_value(row[idx])
            _hydrate_customer_alias_fields(row_dict)
            records.append(row_dict)

        records.sort(
            key=lambda item: str(
                _pick_value(
                    item,
                    ["updatedAt", "UpdatedAt", "updated_at", "createdAt", "CreatedAt", "created_at"],
                    "",
                )
            ),
            reverse=True,
        )
        return records
    except Exception:
        return []


def get_customer_records_by_assignee_email(assignee_email: str) -> list[dict[str, Any]]:
    target = str(assignee_email or "").strip().lower()
    if not target:
        return []

    result: list[dict[str, Any]] = []
    for row in get_all_customer_records():
        if not isinstance(row, dict):
            continue
        matched = False
        for key in CUSTOMER_ASSIGNEE_EMAIL_COLS:
            value = row.get(key)
            if value is None:
                continue
            text = str(value).strip().lower()
            if not text:
                continue
            if text == target or target in text:
                matched = True
                break
        if not matched:
            continue
        result.append(_summary_row(row))

    result.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
    return result


def get_customer_document_content_by_record_id(
    record_id: str,
    field_candidates: list[str],
) -> dict[str, Any] | None:
    record_value = str(record_id or "").strip()
    if not record_value:
        return None

    normalized_fields = [str(value or "").strip() for value in field_candidates if str(value or "").strip()]
    if not normalized_fields:
        return None

    connection = _ensure_connection()
    if not connection:
        return None

    try:
        cursor = connection.cursor()
        table_info = _resolve_customer_table(cursor)
        if not table_info:
            return None

        schema, table, full_table = table_info
        columns = shared_db._list_columns(cursor, schema, table)
        record_col = shared_db._pick_col(columns, CUSTOMER_RECORD_COLS)
        if not record_col:
            return None

        query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({shared_db._qi(record_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s)) LIMIT 1"
        )
        cursor.execute(query, (record_value,))
        row = cursor.fetchone()
        if not row:
            return None

        row_dict = {str(desc[0]): row[idx] for idx, desc in enumerate(cursor.description)}
        lowered_columns = {str(col).strip().lower(): str(col) for col in row_dict.keys()}

        def resolve_actual_column(candidate: str) -> str | None:
            key = str(candidate or "").strip().lower()
            if not key:
                return None
            return lowered_columns.get(key)

        def field_base_candidates(field_name: str) -> list[str]:
            base = str(field_name or "").strip()
            if not base:
                return []
            lowered = base.lower()
            candidates = [base]
            for suffix in ("_filename", "filename", "_contenttype", "contenttype", "_data", "data"):
                if lowered.endswith(suffix):
                    trimmed = base[: -len(suffix)].rstrip("_")
                    if trimmed:
                        candidates.append(trimmed)
            unique: list[str] = []
            seen: set[str] = set()
            for value in candidates:
                marker = value.lower()
                if marker in seen:
                    continue
                seen.add(marker)
                unique.append(value)
            return unique

        for field_name in normalized_fields:
            file_name_hint = ""
            content_type_hint = ""
            attempted_data_cols: list[str] = []

            for base_name in field_base_candidates(field_name):
                data_candidates = [base_name, f"{base_name}_Data", f"{base_name}Data"]
                filename_candidates = [f"{base_name}_FileName", f"{base_name}FileName"]
                content_type_candidates = [f"{base_name}_ContentType", f"{base_name}ContentType"]

                for candidate in filename_candidates:
                    actual = resolve_actual_column(candidate)
                    if not actual:
                        continue
                    raw_name = str(row_dict.get(actual) or "").strip()
                    if raw_name:
                        file_name_hint = raw_name
                        break
                for candidate in content_type_candidates:
                    actual = resolve_actual_column(candidate)
                    if not actual:
                        continue
                    raw_content_type = str(row_dict.get(actual) or "").strip()
                    if raw_content_type:
                        content_type_hint = raw_content_type
                        break

                for candidate in data_candidates:
                    actual = resolve_actual_column(candidate)
                    if not actual or actual in attempted_data_cols:
                        continue
                    attempted_data_cols.append(actual)

            for actual_col in attempted_data_cols:
                file_bytes = _coerce_binary_value(row_dict.get(actual_col))
                if not file_bytes:
                    continue

                file_name = file_name_hint or f"{field_name}.bin"
                content_type = content_type_hint or "application/octet-stream"
                return {
                    "filename": file_name,
                    "content": file_bytes,
                    "content_type": content_type,
                }

        return None
    except Exception:
        return None
