import hashlib
import base64
import binascii
import json
import os
from pathlib import Path
import random
import re
import string
import uuid
from datetime import date, datetime, timedelta, timezone
from typing import Any

import psycopg2
import psycopg2.extensions
from psycopg2 import IntegrityError
from dotenv import load_dotenv

_APP_DIR = Path(__file__).resolve().parent
_PY_BACK_DIR = _APP_DIR.parent
load_dotenv(_APP_DIR / ".env")
load_dotenv(_PY_BACK_DIR / ".env")


def _qi(name: str) -> str:
    """Quote a SQL identifier for PostgreSQL."""
    return '"' + str(name).replace('"', '""') + '"'


def _qt(schema: str, table: str) -> str:
    """Quote schema.table for PostgreSQL."""
    return f"{_qi(schema)}.{_qi(table)}"


def _build_pg_connect_kwargs() -> dict[str, Any]:
    host = (os.getenv("RDS_HOST") or "").strip()
    port = (os.getenv("RDS_PORT") or "5432").strip()
    user = (os.getenv("RDS_USER") or "").strip()
    password = os.getenv("RDS_PASSWORD") or ""
    dbname = (os.getenv("RDS_DB_NAME") or "").strip()
    sslmode = (os.getenv("RDS_SSLMODE") or "require").strip() or "require"
    if not host or not user or not dbname:
        raise RuntimeError(
            "Missing required RDS env vars (RDS_HOST, RDS_USER, RDS_DB_NAME)"
        )
    return {
        "host": host,
        "port": int(port) if str(port).isdigit() else port,
        "user": user,
        "password": password,
        "dbname": dbname,
        "sslmode": sslmode,
    }


conn: psycopg2.extensions.connection | None = None

db_status: dict[str, Any] = {
    "enabled": True,
    "connected": False,
    "source": "RDS_HOST/RDS_PORT/RDS_USER/RDS_DB_NAME",
    "error": None,
}

EMAIL_COLS = ["Email", "email", "EmailID", "emailId", "UserEmail", "userEmail", "Mail", "mail"]
PASSWORD_COLS = ["Password", "password", "UserPassword", "userPassword", "Pass", "pass"]
NAME_COLS = ["Name", "name", "FullName", "fullName", "UserName", "userName"]
ROLE_COLS = ["Role", "role", "UserRole", "userRole"]
BOTNAME_COLS = ["BotName", "botName", "botname", "Bot", "bot"]
ID_COLS = ["Id", "ID", "UserId", "userId"]
VENDOR_TABLE_NAMES = ["VendorMasterData", "Vendor Master Data", "Venodr Master Data"]
ASSIGNED_USER_EMAIL_COLS = [
    "Rubamin Contact Person Email",
    "Rubamin Contact Person",
    "rubaminContactPersonEmail",
    "rubaminContactPerson",
    "invitedByEmail",
    "InvitedByEmail",
    "Invited By Email",
    "buyerEmail",
    "BuyerEmail",
    "contactPersonEmail",
    "ContactPersonEmail",
    "ContactEmailID",
    "contactEmailId",
]
VENDOR_NAME_COLS = ["vendorName", "Vendorname", "VendorName", "Name", "name"]
COMPANY_COLS = ["companyName", "CompanyName", "company", "Company"]
VENDOR_CATEGORY_COLS = ["vendorCategory", "VendorCategory", "Vendor_Category", "category", "Category"]
VENDOR_TYPE_COLS = ["vendorType", "VendorType", "type", "Type"]
APPROVER_STATUS_COLS = ["approverStatus", "ApproverStatus", "ApprovStatus", "inviteStatus", "status", "Status"]
RECORD_COLS = ["recordId", "RecordID", "record", "Record", "Id", "ID"]
CREATED_AT_COLS = ["createdAt", "CreatedAt", "created_on", "CreatedOn", "updatedAt", "UpdatedAt"]
VENDOR_DUPLICATE_PAN_COLS = ["vendorPan", "PANNo", "panNumber", "PANNumber"]
VENDOR_DUPLICATE_CODE_COLS = ["VendorCode", "SAPVendorCode", "BusinessPartnerCode", "vendorCode", "sapVendorCode"]
VENDOR_DUPLICATE_NAME_COLS = ["vendorName", "Vendorname", "VendorName", "Name", "name"]
VENDOR_DUPLICATE_NAME1_COLS = [
    "tradeNameGst",
    "TradeName",
    "tradeNamePan",
    "TradeNamePAN",
    "VendorName1",
    "VendorName2",
    "Name1",
]
VENDOR_DUPLICATE_CITY_COLS = ["city", "City"]
VENDOR_DUPLICATE_DISTRICT_COLS = ["district", "District"]
VENDOR_DUPLICATE_GSTIN_COLS = ["gstNumber", "GSTIN", "gstin"]
VALIDATION_FLAG_COLUMNS = [
    "GST Validate",
    "CIN Validate",
    "PAN Validate",
    "MSME Validate",
    "PF Validate",
    "ESI Validate",
    "Address Validate",
    "Bank Validate",
]
VALIDATION_FLAG_COLUMN_SET = {column.lower() for column in VALIDATION_FLAG_COLUMNS}
DUPLICATE_VENDOR_COLUMN_PAIRS: list[tuple[str, str]] = [
    ("companyName", "CompanyName"),
    ("vendorName", "Vendorname"),
    ("vendorCategory", "Vendor_Category"),
    ("addressLane1", "Address1"),
    ("addressLane2", "Address2"),
    ("addressLane3", "Adddress3"),
    ("country", "Country"),
    ("state", "State"),
    ("district", "District"),
    ("pincode", "Pincode"),
    ("materialDealingsIn", "MaterialDealing"),
    ("serviceProvideFor", "servicedealing"),
    ("bankPaymentMethod", "BankPaymentMethod"),
    ("bankName", "BankName"),
    ("branchName", "BranchName"),
    ("bankAccountNumber", "BankAccount"),
    ("contactPersonName", "ContactPerson"),
    ("contactPersonDesignation", "ContactPersonDesignation"),
    ("contactPersonEmail", "ContactEmailID"),
    ("contactPersonMobile", "ContactPersonMobile"),
    ("alternativePersonName", "AlternativePersonName"),
    ("alternativePersonDesignation", "AlternativePersonDesignation"),
    ("alternativePersonEmail", "AlternativePersonEmail"),
    ("alternativePersonMobile", "AlternativePersonMobile"),
    ("alternativePersonDate", "AlternativePersonDate"),
    ("alternativePersonPlace", "AlternativePersonPlace"),
    ("agreeInformationAccuracy", "AgreeInformationAccuracy"),
    ("Rubamin Contact Person", "rubaminContactPerson"),
    ("Rubamin Contact Person Email", "rubaminContactPersonEmail"),
    ("Rubamin Contact Person Department", "rubaminContactPersonDepartment"),
    ("VendorCode", "SAPVendorCode"),
    ("VendorCode", "BusinessPartnerCode"),
    ("VendorCode", "vendorCode"),
    ("VendorCode", "sapVendorCode"),
    ("VendorCode", "businessPartnerCode"),
]


def _first_non_empty(row: dict[str, Any], keys: list[str], default: Any = None) -> Any:
    for key in keys:
        if key in row and row[key] is not None and str(row[key]).strip() != "":
            return row[key]
    return default


def _normalize_password_value(raw: Any) -> str:
    if raw is None:
        return ""
    if isinstance(raw, bytes):
        try:
            return raw.decode("utf-8", errors="ignore").strip()
        except Exception:
            return str(raw).strip()
    return str(raw).strip()


def _verify_password(input_password: str, stored_password_raw: Any) -> bool:
    input_password = (input_password or "").strip()
    stored_password = _normalize_password_value(stored_password_raw)
    if stored_password.startswith("sha256$"):
        digest = hashlib.sha256(input_password.encode("utf-8")).hexdigest()
        return stored_password == f"sha256${digest}"
    return input_password == stored_password


def _connect() -> psycopg2.extensions.connection:
    return psycopg2.connect(**_build_pg_connect_kwargs())


def initialize_database() -> None:
    global conn
    try:
        conn = _connect()
        cur = conn.cursor()
        cur.execute("SELECT 1")
        cur.fetchone()
        _ensure_vendor_metadata_columns(conn)
        db_status["connected"] = True
        db_status["error"] = None
    except Exception as exc:
        conn = None
        db_status["connected"] = False
        db_status["error"] = str(exc)


def get_db_status() -> dict[str, Any]:
    return dict(db_status)


def _find_table(cursor: psycopg2.extensions.cursor, table_names: list[str]) -> tuple[str, str] | None:
    for table_name in table_names:
        cursor.execute(
            """
            SELECT TABLE_SCHEMA, TABLE_NAME
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_TYPE = 'BASE TABLE'
              AND LOWER(TABLE_NAME) = LOWER(%s)
            ORDER BY CASE WHEN TABLE_SCHEMA='public' THEN 0 ELSE 1 END, TABLE_SCHEMA
            LIMIT 1
            """,
            (table_name,),
        )
        row = cursor.fetchone()
        if row:
            return str(row[0]), str(row[1])
    return None



def _list_columns(cursor: psycopg2.extensions.cursor, schema: str, table: str) -> set[str]:
    cursor.execute(
        """
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s
        """,
        (schema, table),
    )
    return {str(r[0]) for r in cursor.fetchall()}


def _list_column_types(cursor: psycopg2.extensions.cursor, schema: str, table: str) -> dict[str, str]:
    cursor.execute(
        """
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s
        """,
        (schema, table),
    )
    return {str(r[0]).lower(): str(r[1]).lower() for r in cursor.fetchall()}


def _list_non_insertable_columns(cursor: psycopg2.extensions.cursor, schema: str, table: str) -> set[str]:
    cursor.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema = %s AND table_name = %s
          AND (
            is_identity = 'YES'
            OR COALESCE(is_generated, 'NEVER') = 'ALWAYS'
          )
        """,
        (schema, table),
    )
    return {str(r[0]).lower() for r in cursor.fetchall()}



def _list_required_non_nullable_columns(
    cursor: psycopg2.extensions.cursor,
    schema: str,
    table: str,
) -> list[dict[str, str]]:
    cursor.execute(
        """
        SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, COLUMN_DEFAULT
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s
        ORDER BY ORDINAL_POSITION
        """,
        (schema, table),
    )

    required: list[dict[str, str]] = []
    for row in cursor.fetchall():
        column_name = str(row[0])
        data_type = str(row[1] or "").lower()
        is_nullable = str(row[2] or "").upper()
        has_default = row[3] is not None
        if is_nullable == "YES" or has_default:
            continue
        required.append({"name": column_name, "data_type": data_type})
    return required


def _fallback_required_user_column_value(
    column_name: str,
    data_type: str,
    *,
    name: str,
    email: str,
    role: str,
    bot_name: str | None = None,
    password_value: str,
    now_value: str,
) -> Any:
    column_key = str(column_name or "").strip().lower()
    sql_type = str(data_type or "").strip().lower()

    if "password" in column_key or column_key in {"pass", "pwd"}:
        return password_value
    if "email" in column_key or "mail" in column_key:
        return email
    if "role" in column_key:
        return role
    if "botname" in column_key:
        return bot_name or name or email or "BOT"
    if "name" in column_key:
        return name or email or "User"
    if "created" in column_key or "updated" in column_key or "date" in column_key or "time" in column_key:
        return now_value

    if sql_type in {"bit", "boolean", "bool"}:
        return False
    if sql_type in {"int", "bigint", "smallint", "tinyint"}:
        return 0
    if sql_type in {"decimal", "numeric", "float", "real", "money", "smallmoney"}:
        return 0
    if sql_type in {"uniqueidentifier", "uuid"}:
        return str(uuid.uuid4())
    if sql_type in {"date", "datetime", "datetime2", "smalldatetime", "datetimeoffset", "time"}:
        return now_value

    return "N/A"


def _pick_col(columns: set[str], candidates: list[str]) -> str | None:
    lower_map = {c.lower(): c for c in columns}
    for candidate in candidates:
        if candidate.lower() in lower_map:
            return lower_map[candidate.lower()]
    return None


def _resolve_column_name(columns: set[str], candidate: str) -> str | None:
    target = str(candidate or "").strip()
    if not target:
        return None
    if target in columns:
        return target
    for column_name in columns:
        if column_name.lower() == target.lower():
            return column_name
    return None


def _normalize_validation_flag_value(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, (int, float)):
        return "True" if value != 0 else "False"

    text = str(value).strip()
    if text == "":
        return None

    lowered = text.lower()
    if lowered in {"1", "true", "yes", "y", "on"}:
        return "True"
    if lowered in {"0", "false", "no", "n", "off"}:
        return "False"
    return "False"


def _ensure_validation_flag_columns_as_text(
    cursor: psycopg2.extensions.cursor,
    schema: str,
    table: str,
) -> None:
    columns = _list_columns(cursor, schema, table)
    col_map = {column.lower(): column for column in columns}
    col_types = _list_column_types(cursor, schema, table)

    for expected_column in VALIDATION_FLAG_COLUMNS:
        key = expected_column.lower()
        actual_col = col_map.get(key)
        if not actual_col:
            continue

        sql_type = col_types.get(actual_col.lower(), "")
        final_sql_type = sql_type
        if sql_type not in {"character varying", "varchar", "text"}:
            try:
                cursor.execute(
                    f"ALTER TABLE {_qt(schema, table)} ALTER COLUMN {_qi(actual_col)} TYPE VARCHAR(5) USING {_qi(actual_col)}::text"
                )
                final_sql_type = "varchar"
            except Exception:
                # If column cannot be altered (constraints, permissions), keep existing type.
                final_sql_type = sql_type

        if final_sql_type in {"character varying", "varchar", "text"}:
            cursor.execute(
                f"""
                UPDATE {_qt(schema, table)}
                SET {_qi(actual_col)} = CASE
                    WHEN {_qi(actual_col)} IS NULL OR TRIM(CAST({_qi(actual_col)} AS VARCHAR(50))) = '' THEN NULL
                    WHEN LOWER(TRIM(CAST({_qi(actual_col)} AS VARCHAR(50)))) IN ('1', 'true', 'yes', 'y', 'on') THEN 'True'
                    WHEN LOWER(TRIM(CAST({_qi(actual_col)} AS VARCHAR(50)))) IN ('0', 'false', 'no', 'n', 'off') THEN 'False'
                    ELSE 'False'
                END
                WHERE {_qi(actual_col)} IS NOT NULL
                """
            )
        elif final_sql_type in {"bit", "boolean", "bool"}:
            cursor.execute(
                f"""
                UPDATE {_qt(schema, table)}
                SET {_qi(actual_col)} = CASE
                    WHEN {_qi(actual_col)} IS NULL THEN NULL
                    WHEN LOWER(TRIM(CAST({_qi(actual_col)} AS VARCHAR(50)))) IN ('1', 'true', 'yes', 'y', 'on') THEN 1
                    ELSE 0
                END
                WHERE {_qi(actual_col)} IS NOT NULL
                """
            )


def _ensure_vendor_metadata_columns(connection: psycopg2.extensions.connection) -> None:
    try:
        cursor = connection.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return

        schema, table = vendor_table
        existing_columns = {col.lower() for col in _list_columns(cursor, schema, table)}
        required_columns = [
            ("VendorType", "VARCHAR(50) NULL"),
            ("IncoTerms", "VARCHAR(200) NULL"),
            ("PaymentTerms", "VARCHAR(500) NULL"),
            ("WithholdingTax", "TEXT NULL"),
            ("TurnoverLimit", "VARCHAR(50) NULL"),
            ("RubaminApproverHod", "VARCHAR(320) NULL"),
            ("Rubamin Contact Person", "VARCHAR(500) NULL"),
            ("Rubamin Contact Person Email", "VARCHAR(320) NULL"),
            ("Rubamin Contact Person Department", "VARCHAR(200) NULL"),
            ("ApproverStatus", "VARCHAR(50) NULL"),
            ("ReviewDecision", "VARCHAR(50) NULL"),
            ("ReviewRemarks", "TEXT NULL"),
            ("ReviewedByRole", "VARCHAR(100) NULL"),
            ("VendorCode", "VARCHAR(50) NULL"),
            ("SAPCodeGeneratedAt", "VARCHAR(50) NULL"),
            ("TaxpayerTypeGST", "VARCHAR(200) NULL"),
            ("GST Validate", "VARCHAR(5) NULL"),
            ("CIN Validate", "VARCHAR(5) NULL"),
            ("PAN Validate", "VARCHAR(5) NULL"),
            ("MSME Validate", "VARCHAR(5) NULL"),
            ("PF Validate", "VARCHAR(5) NULL"),
            ("ESI Validate", "VARCHAR(5) NULL"),
            ("Address Validate", "VARCHAR(5) NULL"),
            ("Bank Validate", "VARCHAR(5) NULL"),
            ("Validate By", "VARCHAR(320) NULL"),
        ]

        for column_name, column_sql in required_columns:
            if column_name.lower() in existing_columns:
                continue
            cursor.execute(f"ALTER TABLE {_qt(schema, table)} ADD COLUMN {_qi(column_name)} {column_sql}")
            existing_columns.add(column_name.lower())

        _ensure_validation_flag_columns_as_text(cursor, schema, table)
        columns_after = _list_columns(cursor, schema, table)
        rubamin_name_col = _resolve_column_name(columns_after, "Rubamin Contact Person")
        rubamin_email_col = _resolve_column_name(columns_after, "Rubamin Contact Person Email")
        rubamin_dept_col = _resolve_column_name(columns_after, "Rubamin Contact Person Department")
        department_col = _resolve_column_name(columns_after, "department")

        if rubamin_name_col and rubamin_email_col:
            cursor.execute(
                f"""
                UPDATE {_qt(schema, table)}
                SET {_qi(rubamin_email_col)} = TRIM(CAST({_qi(rubamin_name_col)} AS VARCHAR(320)))
                WHERE {_qi(rubamin_name_col)} IS NOT NULL
                  AND STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '@') > 0
                  AND STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '<') = 0
                  AND (
                    {_qi(rubamin_email_col)} IS NULL
                    OR TRIM(CAST({_qi(rubamin_email_col)} AS VARCHAR(320))) = ''
                  )
                """
            )
            cursor.execute(
                f"""
                UPDATE {_qt(schema, table)}
                SET {_qi(rubamin_email_col)} = TRIM(
                    SUBSTRING(
                        CAST({_qi(rubamin_name_col)} AS VARCHAR(500)),
                        STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '<') + 1,
                        STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '>') - STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '<') - 1
                    )
                )
                WHERE {_qi(rubamin_name_col)} IS NOT NULL
                  AND STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '<') > 0
                  AND STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '>') > STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '<')
                  AND STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '@') > 0
                  AND (
                    {_qi(rubamin_email_col)} IS NULL
                    OR TRIM(CAST({_qi(rubamin_email_col)} AS VARCHAR(320))) = ''
                  )
                """
            )
            cursor.execute(
                f"""
                UPDATE {_qt(schema, table)}
                SET {_qi(rubamin_name_col)} = NULLIF(
                    TRIM(
                        CASE
                            WHEN STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '<') > 0
                            THEN LEFT(
                                CAST({_qi(rubamin_name_col)} AS VARCHAR(500)),
                                STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '<') - 1
                            )
                            WHEN STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '@') > 1
                            THEN LEFT(
                                CAST({_qi(rubamin_name_col)} AS VARCHAR(500)),
                                STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '@') - 1
                            )
                            ELSE CAST({_qi(rubamin_name_col)} AS VARCHAR(500))
                        END
                    ),
                    ''
                )
                WHERE {_qi(rubamin_name_col)} IS NOT NULL
                  AND STRPOS(CAST({_qi(rubamin_name_col)} AS VARCHAR(500)), '@') > 0
                """
            )

        if rubamin_dept_col and department_col and rubamin_dept_col.lower() != department_col.lower():
            cursor.execute(
                f"""
                UPDATE {_qt(schema, table)}
                SET {_qi(rubamin_dept_col)} = CAST({_qi(department_col)} AS VARCHAR(200))
                WHERE {_qi(department_col)} IS NOT NULL
                  AND (
                    {_qi(rubamin_dept_col)} IS NULL
                    OR TRIM(CAST({_qi(rubamin_dept_col)} AS VARCHAR(200))) = ''
                  )
                """
            )

        _cleanup_duplicate_vendor_columns(cursor, schema, table)
        connection.commit()
    except Exception:
        try:
            connection.rollback()
        except Exception:
            pass


def _cleanup_duplicate_vendor_columns(
    cursor: psycopg2.extensions.cursor,
    schema: str,
    table: str,
) -> list[str]:
    dropped_columns: list[str] = []
    columns = _list_columns(cursor, schema, table)

    for canonical_name, duplicate_name in DUPLICATE_VENDOR_COLUMN_PAIRS:
        canonical_col = _resolve_column_name(columns, canonical_name)
        duplicate_col = _resolve_column_name(columns, duplicate_name)
        if not canonical_col or not duplicate_col:
            continue
        if canonical_col == duplicate_col:
            continue

        try:
            cursor.execute(
                f"""
                UPDATE {_qt(schema, table)}
                SET {_qi(canonical_col)} = {_qi(duplicate_col)}
                WHERE {_qi(duplicate_col)} IS NOT NULL
                  AND (
                    {_qi(canonical_col)} IS NULL
                    OR TRIM(CAST({_qi(canonical_col)} AS TEXT)) = ''
                  )
                """
            )
            cursor.execute(f"ALTER TABLE {_qt(schema, table)} DROP COLUMN {_qi(duplicate_col)}")
            dropped_columns.append(duplicate_col)
            columns.discard(duplicate_col)
        except Exception:
            # Keep startup stable even if a column cannot be dropped due constraints.
            continue

    return dropped_columns


def cleanup_vendor_duplicate_columns() -> dict[str, Any]:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        return {
            "success": False,
            "droppedColumns": [],
            "error": db_status.get("error") or "Database connection failed",
        }

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return {"success": False, "droppedColumns": [], "error": "Vendor master table not found"}

        schema, table = vendor_table
        dropped = _cleanup_duplicate_vendor_columns(cursor, schema, table)
        conn.commit()
        return {
            "success": True,
            "schema": schema,
            "table": table,
            "droppedColumns": dropped,
        }
    except Exception as exc:
        try:
            conn.rollback()
        except Exception:
            pass
        return {
            "success": False,
            "droppedColumns": [],
            "error": str(exc),
        }


def ensure_vendor_columns(column_names: list[str]) -> list[str]:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        raise RuntimeError(db_status.get("error") or "Database connection failed")

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            raise RuntimeError("Vendor master table not found")

        schema, table = vendor_table
        existing_columns = _list_columns(cursor, schema, table)
        existing_lower = {column.lower() for column in existing_columns}
        added_columns: list[str] = []

        for raw_name in column_names:
            column_name = str(raw_name or "").strip()
            if not column_name:
                continue
            if '"' in column_name:
                continue

            lowered = column_name.lower()
            if lowered in existing_lower:
                continue

            column_sql = "TEXT NULL"
            if lowered.endswith("_data"):
                column_sql = "BYTEA NULL"
            elif lowered.endswith("_filename"):
                column_sql = "VARCHAR(500) NULL"
            elif lowered.endswith("_contenttype"):
                column_sql = "VARCHAR(200) NULL"

            cursor.execute(
                f"ALTER TABLE {_qt(schema, table)} ADD COLUMN {_qi(column_name)} {column_sql}"
            )
            existing_lower.add(lowered)
            added_columns.append(column_name)

        if added_columns:
            conn.commit()
        return added_columns
    except Exception as exc:
        try:
            conn.rollback()
        except Exception:
            pass
        raise RuntimeError(f"Unable to ensure vendor columns: {exc}") from exc


def _generate_record_id(length: int = 24) -> str:
    chars = string.ascii_uppercase + string.ascii_lowercase + string.digits
    return "".join(random.choice(chars) for _ in range(length))


def _coerce_sql_value(value: Any, sql_type: str) -> Any:
    if value is None:
        return None

    if isinstance(value, str):
        value = value.strip()
        if value == "":
            return None

    if sql_type in {"int", "bigint", "smallint", "tinyint"}:
        if isinstance(value, int):
            return value
        text = str(value)
        # Supports values like "2025-26" by taking the first integer token.
        match = re.search(r"-?\d+", text)
        if not match:
            return None
        return int(match.group(0))

    if sql_type in {"decimal", "numeric", "float", "real", "money", "smallmoney"}:
        if isinstance(value, (int, float)):
            return value
        text = str(value).replace(",", "").strip()
        match = re.search(r"-?\d+(?:\.\d+)?", text)
        if not match:
            return None
        return float(match.group(0))

    if sql_type in {"bit", "boolean", "bool"}:
        if isinstance(value, bool):
            return value
        text = str(value).strip().lower()
        return text in {"1", "true", "yes", "y", "on"}

    return value


def authenticate_db_user(login_id: str, password: str) -> dict[str, Any] | None:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        return None

    try:
        cursor = conn.cursor()

        users_table = _find_table(cursor, ["Users"])
        if not users_table:
            db_status["error"] = "Users table not found"
            return None

        schema, table = users_table
        full_table = _qt(schema, table)

        cols = _list_columns(cursor, schema, table)
        email_col = _pick_col(cols, EMAIL_COLS)
        name_col = _pick_col(cols, NAME_COLS)
        pass_col = _pick_col(cols, PASSWORD_COLS)
        if (not email_col and not name_col) or not pass_col:
            db_status["error"] = f"Required columns not found in Users table. Found: {sorted(cols)}"
            return None

        if email_col and name_col:
            query = (
                f"SELECT * FROM {full_table} "
                f"WHERE LOWER(TRIM(CAST({_qi(email_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s)) "
                f"OR LOWER(TRIM(CAST({_qi(name_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s)) LIMIT 1"
            )
            cursor.execute(query, (login_id, login_id))
        elif email_col:
            query = (
                f"SELECT * FROM {full_table} "
                f"WHERE LOWER(TRIM(CAST({_qi(email_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s)) LIMIT 1"
            )
            cursor.execute(query, (login_id,))
        else:
            query = (
                f"SELECT * FROM {full_table} "
                f"WHERE LOWER(TRIM(CAST({_qi(name_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s)) LIMIT 1"
            )
            cursor.execute(query, (login_id,))
        row = cursor.fetchone()
        if not row:
            return None

        row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}

        stored_password = _first_non_empty(row_dict, [pass_col], "")
        if not _verify_password(password, stored_password):
            return None

        user_id = _first_non_empty(row_dict, ID_COLS, None)
        name = str(_first_non_empty(row_dict, NAME_COLS, "SQL User"))
        role = str(_first_non_empty(row_dict, ROLE_COLS, "Vendor"))
        user_email = str(_first_non_empty(row_dict, [email_col] if email_col else [name_col], login_id))

        try:
            bot_col = _pick_col(cols, BOTNAME_COLS)
        except Exception:
            bot_col = None
        bot_name = ""
        if bot_col:
            bot_name = str(_first_non_empty(row_dict, [bot_col], "Vendor Bot")).strip()
        bot_name = bot_name or "Vendor Bot"

        return {
            "id": user_id,
            "name": name,
            "email": user_email,
            "role": role,
            "botName": bot_name,
        }
    except Exception as exc:
        db_status["connected"] = False
        db_status["error"] = str(exc)
        conn = None
        return None


def update_db_user_password_by_email(email: str, new_password: str) -> bool:
    global conn

    email_value = str(email or "").strip()
    password_value = str(new_password or "").strip()
    if not email_value or not password_value:
        return False

    if not conn:
        initialize_database()
    if not conn:
        return False

    try:
        cursor = conn.cursor()
        users_table = _find_table(cursor, ["Users"])
        if not users_table:
            return False

        schema, table = users_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)
        blocked_columns = _list_non_insertable_columns(cursor, schema, table)

        email_col = _pick_col(cols, EMAIL_COLS)
        pass_col = _pick_col(cols, PASSWORD_COLS)
        if not email_col or not pass_col:
            return False
        if pass_col.lower() in blocked_columns:
            raise RuntimeError("Password column is not updateable in Users table")

        query = (
            f"UPDATE {full_table} "
            f"SET {_qi(pass_col)} = %s "
            f"WHERE LOWER(TRIM(CAST({_qi(email_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s))"
        )
        cursor.execute(query, (_hash_password_for_storage(password_value), email_value))
        affected_rows = cursor.rowcount
        conn.commit()
        return affected_rows > 0
    except Exception as exc:
        try:
            conn.rollback()
        except Exception:
            pass
        raise RuntimeError(f"Unable to update password in SQL Users table: {exc}") from exc


def update_db_user_by_email(
    current_email: str,
    *,
    name: str,
    email: str,
    role: str,
    bot_name: str,
) -> dict[str, Any] | None:
    global conn

    current_email_value = str(current_email or "").strip()
    next_email_value = str(email or "").strip()
    next_name_value = str(name or "").strip()
    next_role_value = str(role or "").strip()
    next_bot_value = str(bot_name or "").strip()
    if not current_email_value:
        return None
    if not next_email_value or not next_name_value or not next_role_value:
        raise RuntimeError("Name, email and role are required")

    if not conn:
        initialize_database()
    if not conn:
        raise RuntimeError(db_status.get("error") or "Database connection failed")

    try:
        cursor = conn.cursor()
        users_table = _find_table(cursor, ["Users"])
        if not users_table:
            raise RuntimeError("Users table not found")

        schema, table = users_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)
        blocked_columns = _list_non_insertable_columns(cursor, schema, table)

        email_col = _pick_col(cols, EMAIL_COLS)
        name_col = _pick_col(cols, NAME_COLS)
        role_col = _pick_col(cols, ROLE_COLS)
        bot_col = _pick_col(cols, BOTNAME_COLS)
        id_col = _pick_col(cols, ID_COLS)
        if not email_col:
            raise RuntimeError("Email column not found in Users table")

        existing_query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({_qi(email_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s)) LIMIT 1"
        )
        cursor.execute(existing_query, (current_email_value,))
        existing_row = cursor.fetchone()
        if not existing_row:
            return None

        set_parts: list[str] = []
        params: list[Any] = []

        if name_col and name_col.lower() not in blocked_columns:
            set_parts.append(f"{_qi(name_col)} = %s")
            params.append(next_name_value)
        if email_col.lower() not in blocked_columns:
            set_parts.append(f"{_qi(email_col)} = %s")
            params.append(next_email_value)
        if role_col and role_col.lower() not in blocked_columns:
            set_parts.append(f"{_qi(role_col)} = %s")
            params.append(next_role_value)
        if bot_col and bot_col.lower() not in blocked_columns:
            set_parts.append(f"{_qi(bot_col)} = %s")
            params.append(next_bot_value or "Vendor Bot")

        if not set_parts:
            raise RuntimeError("No updateable user columns found")

        update_query = (
            f"UPDATE {full_table} "
            f"SET {', '.join(set_parts)} "
            f"WHERE LOWER(TRIM(CAST({_qi(email_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s))"
        )
        params.append(current_email_value)
        cursor.execute(update_query, params)
        if cursor.rowcount <= 0:
            conn.rollback()
            return None

        conn.commit()

        cursor.execute(existing_query, (next_email_value,))
        updated_row = cursor.fetchone()
        if not updated_row:
            return {
                "id": None,
                "name": next_name_value,
                "email": next_email_value,
                "role": next_role_value,
                "botName": next_bot_value or "Vendor Bot",
            }

        row_dict = {desc[0]: updated_row[idx] for idx, desc in enumerate(cursor.description)}
        mapped = _map_user_row(row_dict, id_col, name_col, email_col, role_col, bot_col=bot_col)
        if mapped:
            return mapped

        return {
            "id": None,
            "name": next_name_value,
            "email": next_email_value,
            "role": next_role_value,
            "botName": next_bot_value or "Vendor Bot",
        }
    except Exception as exc:
        try:
            conn.rollback()
        except Exception:
            pass
        raise RuntimeError(f"Unable to update user in SQL Users table: {exc}") from exc


def get_db_users_by_role(role_value: str) -> list[dict[str, Any]]:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        return []

    try:
        cursor = conn.cursor()
        users_table = _find_table(cursor, ["Users"])
        if not users_table:
            return []

        schema, table = users_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)

        role_col = _pick_col(cols, ROLE_COLS)
        email_col = _pick_col(cols, EMAIL_COLS)
        name_col = _pick_col(cols, NAME_COLS)
        id_col = _pick_col(cols, ID_COLS)
        if not role_col or not email_col:
            return []

        query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({_qi(role_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s))"
        )
        cursor.execute(query, (role_value,))
        rows = cursor.fetchall()

        result: list[dict[str, Any]] = []
        for row in rows:
            row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}
            email = str(_first_non_empty(row_dict, [email_col], "")).strip()
            if not email:
                continue

            result.append(
                {
                    "id": _first_non_empty(row_dict, [id_col] if id_col else [], None),
                    "name": str(_first_non_empty(row_dict, [name_col] if name_col else [email_col], email)).strip(),
                    "email": email,
                    "role": str(_first_non_empty(row_dict, [role_col], role_value)).strip(),
                }
            )

        return result
    except Exception:
        return []


def _hash_password_for_storage(password: str) -> str:
    return str(password or "").strip()


def _map_user_row(
    row_dict: dict[str, Any],
    id_col: str | None,
    name_col: str | None,
    email_col: str | None,
    role_col: str | None,
    bot_col: str | None = None,
) -> dict[str, Any] | None:
    email_candidates = [email_col] if email_col else []
    email = str(_first_non_empty(row_dict, email_candidates, "")).strip()
    if not email:
        return None

    name_candidates = [name_col] if name_col else []
    role_candidates = [role_col] if role_col else []
    id_candidates = [id_col] if id_col else []
    bot_candidates = [bot_col] if bot_col else []

    return {
        "id": _first_non_empty(row_dict, id_candidates, None),
        "name": str(_first_non_empty(row_dict, name_candidates or email_candidates, email)).strip(),
        "email": email,
        "role": str(_first_non_empty(row_dict, role_candidates, "User")).strip(),
        "botName": str(_first_non_empty(row_dict, bot_candidates, "Vendor Bot")).strip(),
    }


def get_db_users() -> list[dict[str, Any]]:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        return []

    try:
        cursor = conn.cursor()
        users_table = _find_table(cursor, ["Users"])
        if not users_table:
            return []

        schema, table = users_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)

        email_col = _pick_col(cols, EMAIL_COLS)
        name_col = _pick_col(cols, NAME_COLS)
        role_col = _pick_col(cols, ROLE_COLS)
        bot_col = _pick_col(cols, BOTNAME_COLS)
        id_col = _pick_col(cols, ID_COLS)
        if not email_col:
            return []

        cursor.execute(f"SELECT * FROM {full_table}")
        rows = cursor.fetchall()

        result: list[dict[str, Any]] = []
        for row in rows:
            row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}
            mapped = _map_user_row(row_dict, id_col, name_col, email_col, role_col, bot_col=bot_col)
            if mapped:
                result.append(mapped)

        result.sort(key=lambda item: (str(item.get("name") or "").lower(), str(item.get("email") or "").lower()))
        return result
    except Exception:
        return []


def get_vendor_records_by_assignee_email(assignee_email: str) -> list[dict[str, Any]]:
    global conn

    email_value = str(assignee_email or "").strip()
    if not email_value:
        return []

    if not conn:
        initialize_database()
    if not conn:
        return []

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return []

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)
        col_map = {c.lower(): c for c in cols}

        assignee_cols: list[str] = []
        for candidate in ASSIGNED_USER_EMAIL_COLS:
            actual_col = col_map.get(candidate.lower())
            if actual_col and actual_col not in assignee_cols:
                assignee_cols.append(actual_col)
        if not assignee_cols:
            return []

        record_col = _pick_col(cols, RECORD_COLS)
        company_col = _pick_col(cols, COMPANY_COLS)
        category_col = _pick_col(cols, VENDOR_CATEGORY_COLS)
        approver_status_col = _pick_col(cols, APPROVER_STATUS_COLS)
        name_col = _pick_col(cols, VENDOR_NAME_COLS)
        type_col = _pick_col(cols, VENDOR_TYPE_COLS)
        created_at_col = _pick_col(cols, CREATED_AT_COLS)

        where_parts = []
        params: list[str] = []
        for actual_col in assignee_cols:
            normalized_expr = f"LOWER(TRIM(CAST({_qi(actual_col)} AS VARCHAR(1000))))"
            where_parts.append(
                f"({normalized_expr} = LOWER(TRIM(%s)) OR {normalized_expr} LIKE '%' || LOWER(TRIM(%s)) || '%')"
            )
            params.extend([email_value, email_value])

        query = f"SELECT * FROM {full_table} WHERE {' OR '.join(where_parts)}"
        cursor.execute(query, params)
        db_rows = cursor.fetchall()

        result: list[dict[str, Any]] = []
        seen_records: set[str] = set()
        for row in db_rows:
            row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}

            def pick(column_name: str | None, default: str = "") -> str:
                if not column_name:
                    return default
                value = row_dict.get(column_name)
                if value is None:
                    return default
                return str(value).strip()

            record_value = pick(record_col)
            record_key = record_value.lower()
            if record_key and record_key in seen_records:
                continue
            if record_key:
                seen_records.add(record_key)

            result.append(
                {
                    "record": record_value,
                    "company": pick(company_col),
                    "category": pick(category_col),
                    "approverStatus": pick(approver_status_col),
                    "name": pick(name_col),
                    "type": pick(type_col),
                    "createdAt": pick(created_at_col),
                }
            )

        result.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
        return result
    except Exception:
        return []


def get_vendor_records_by_approver_status(approver_status: str) -> list[dict[str, Any]]:
    global conn

    status_value = str(approver_status or "").strip()
    if not status_value:
        return []

    if not conn:
        initialize_database()
    if not conn:
        return []

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return []

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)

        status_col = _pick_col(cols, APPROVER_STATUS_COLS)
        if not status_col:
            return []

        record_col = _pick_col(cols, RECORD_COLS)
        company_col = _pick_col(cols, COMPANY_COLS)
        category_col = _pick_col(cols, VENDOR_CATEGORY_COLS)
        name_col = _pick_col(cols, VENDOR_NAME_COLS)
        type_col = _pick_col(cols, VENDOR_TYPE_COLS)
        created_at_col = _pick_col(cols, CREATED_AT_COLS)

        query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({_qi(status_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s))"
        )
        cursor.execute(query, (status_value,))
        db_rows = cursor.fetchall()

        result: list[dict[str, Any]] = []
        for row in db_rows:
            row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}

            def pick(column_name: str | None, default: str = "") -> str:
                if not column_name:
                    return default
                value = row_dict.get(column_name)
                if value is None:
                    return default
                return str(value).strip()

            result.append(
                {
                    "record": pick(record_col),
                    "company": pick(company_col),
                    "category": pick(category_col),
                    "approverStatus": pick(status_col),
                    "name": pick(name_col),
                    "type": pick(type_col),
                    "createdAt": pick(created_at_col),
                }
            )

        result.sort(key=lambda item: str(item.get("createdAt") or ""), reverse=True)
        return result
    except Exception:
        return []


def get_all_vendor_records() -> list[dict[str, Any]]:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        return []

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return []

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cursor.execute(f"SELECT * FROM {full_table}")
        db_rows = cursor.fetchall()
        description = cursor.description

        result: list[dict[str, Any]] = []
        for row in db_rows:
            row_dict: dict[str, Any] = {}
            for idx, desc in enumerate(description):
                row_dict[str(desc[0])] = _to_json_safe_value(row[idx])
            result.append(row_dict)

        def sort_key(item: dict[str, Any]) -> str:
            for key in ("updatedAt", "UpdatedAt", "createdAt", "CreatedAt", "created_on", "CreatedOn"):
                value = str(item.get(key) or "").strip()
                if value:
                    return value
            return ""

        result.sort(key=sort_key, reverse=True)
        return result
    except Exception:
        return []


def get_vendor_duplicate_records_by_pan(pan_number: str) -> list[dict[str, Any]]:
    global conn

    pan_value = str(pan_number or "").strip().upper()
    if not pan_value:
        return []

    if not conn:
        initialize_database()
    if not conn:
        return []

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return []

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)

        pan_cols: list[str] = []
        for candidate in VENDOR_DUPLICATE_PAN_COLS:
            actual_col = _resolve_column_name_loose(cols, candidate)
            if actual_col and actual_col not in pan_cols:
                pan_cols.append(actual_col)
        if not pan_cols:
            return []

        updated_at_col = _resolve_column_name_loose(cols, "updatedAt")
        created_at_col = _resolve_column_name_loose(cols, "createdAt") or _resolve_column_name_loose(cols, "CreatedOn")

        where_parts: list[str] = []
        params: list[str] = []
        for actual_col in pan_cols:
            where_parts.append(
                f"UPPER(TRIM(CAST({_qi(actual_col)} AS VARCHAR(200)))) = UPPER(TRIM(%s))"
            )
            params.append(pan_value)

        order_by_parts: list[str] = []
        if updated_at_col:
            order_by_parts.append(f"{_qi(updated_at_col)} DESC")
        if created_at_col and created_at_col != updated_at_col:
            order_by_parts.append(f"{_qi(created_at_col)} DESC")

        query = f"SELECT * FROM {full_table} WHERE {' OR '.join(where_parts)}"
        if order_by_parts:
            query += f" ORDER BY {', '.join(order_by_parts)}"

        cursor.execute(query, params)
        db_rows = cursor.fetchall()

        result: list[dict[str, Any]] = []
        seen_keys: set[str] = set()
        for row in db_rows:
            row_dict: dict[str, Any] = {}
            for idx, desc in enumerate(cursor.description):
                row_dict[str(desc[0])] = _to_json_safe_value(row[idx])
            row_columns = set(row_dict.keys())

            def pick(candidates: list[str], default: str = "") -> str:
                for candidate in candidates:
                    actual_col = _resolve_column_name_loose(row_columns, candidate)
                    if not actual_col:
                        continue
                    value = row_dict.get(actual_col)
                    if value is None:
                        continue
                    text = str(value).strip()
                    if text:
                        return text
                return default

            record_id = pick(RECORD_COLS)
            vendor_code = pick(VENDOR_DUPLICATE_CODE_COLS)
            name = pick(VENDOR_DUPLICATE_NAME_COLS)
            name1 = pick(VENDOR_DUPLICATE_NAME1_COLS)
            city = pick(VENDOR_DUPLICATE_CITY_COLS)
            district = pick(VENDOR_DUPLICATE_DISTRICT_COLS)
            matched_pan = pick(VENDOR_DUPLICATE_PAN_COLS, pan_value)
            gstin = pick(VENDOR_DUPLICATE_GSTIN_COLS)
            updated_at = pick(["updatedAt", "UpdatedAt", "createdAt", "CreatedAt", "created_on", "CreatedOn"])

            dedupe_key = "|".join(
                [
                    record_id.lower(),
                    vendor_code.lower(),
                    name.lower(),
                    matched_pan.lower(),
                    gstin.lower(),
                ]
            )
            if dedupe_key in seen_keys:
                continue
            seen_keys.add(dedupe_key)

            result.append(
                {
                    "recordId": record_id,
                    "vendorCode": vendor_code,
                    "name": name,
                    "name1": name1,
                    "city": city,
                    "district": district,
                    "pan": matched_pan,
                    "gstin": gstin,
                    "updatedAt": updated_at,
                }
            )

        return result
    except Exception:
        return []


def _to_json_safe_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (bytes, bytearray, memoryview)):
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, uuid.UUID):
        return str(value)
    return value


def get_vendor_record_by_record_id(record_id: str) -> dict[str, Any] | None:
    global conn

    record_value = str(record_id or "").strip()
    if not record_value:
        return None

    if not conn:
        initialize_database()
    if not conn:
        return None

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return None

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)

        record_col = _pick_col(cols, RECORD_COLS)
        if not record_col:
            return None

        query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({_qi(record_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s)) LIMIT 1"
        )
        cursor.execute(query, (record_value,))
        row = cursor.fetchone()
        if not row:
            return None

        row_dict: dict[str, Any] = {}
        for idx, desc in enumerate(cursor.description):
            key = str(desc[0])
            row_dict[key] = _to_json_safe_value(row[idx])

        return row_dict
    except Exception:
        return None


def _document_column_base_name(column_name: str) -> str:
    value = str(column_name or "")
    lowered = value.lower()
    if lowered.endswith("_filename"):
        return value[: -len("_FileName")]
    if lowered.endswith("filename"):
        return value[: -len("FileName")]
    return value


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


def _normalize_column_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _resolve_column_name_loose(columns: set[str], candidate: str) -> str | None:
    exact = _resolve_column_name(columns, candidate)
    if exact:
        return exact

    target = _normalize_column_key(candidate)
    if not target:
        return None

    for column_name in columns:
        if _normalize_column_key(column_name) == target:
            return column_name
    return None


def _extract_binary_from_text(raw_text: str) -> dict[str, Any] | None:
    text = str(raw_text or "").strip()
    if not text:
        return None

    lowered = text.lower()
    if lowered.startswith("data:") and ";base64," in lowered:
        try:
            header, encoded = text.split(",", 1)
            media_type = "application/octet-stream"
            if ":" in header and ";" in header:
                media_type = header.split(":", 1)[1].split(";", 1)[0] or media_type
            content = base64.b64decode(encoded, validate=False)
            return {
                "content": content,
                "content_type": media_type,
            }
        except Exception:
            return None

    if lowered.startswith("0x") and len(text) > 2:
        try:
            content = bytes.fromhex(text[2:])
            return {
                "content": content,
                "content_type": "application/octet-stream",
            }
        except ValueError:
            return None

    if not ((text.startswith("{") and text.endswith("}")) or (text.startswith("[") and text.endswith("]"))):
        return None

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        return None

    if isinstance(parsed, list):
        for item in parsed:
            if isinstance(item, (dict, list, str)):
                found = _extract_binary_from_text(json.dumps(item) if isinstance(item, (dict, list)) else item)
                if found:
                    return found
        return None

    if not isinstance(parsed, dict):
        return None

    filename_candidates = (
        "fileName",
        "filename",
        "name",
        "displayName",
        "title",
    )
    content_type_candidates = (
        "$content-type",
        "contentType",
        "mimeType",
        "mediaType",
        "type",
    )
    content_candidates = (
        "$content",
        "content",
        "contentBytes",
        "fileContent",
        "value",
        "data",
    )

    filename = ""
    for key in filename_candidates:
        value = Path(str(parsed.get(key) or "")).name.strip()
        if value:
            filename = value
            break

    content_type = "application/octet-stream"
    for key in content_type_candidates:
        value = str(parsed.get(key) or "").strip()
        if value:
            content_type = value
            break

    for key in content_candidates:
        raw_value = parsed.get(key)
        if raw_value is None:
            continue
        if isinstance(raw_value, (bytes, bytearray, memoryview)):
            return {
                "filename": filename,
                "content": _coerce_binary_value(raw_value) or b"",
                "content_type": content_type,
            }
        value_text = str(raw_value).strip()
        if not value_text:
            continue

        nested = _extract_binary_from_text(value_text)
        if nested:
            if filename and not nested.get("filename"):
                nested["filename"] = filename
            if content_type and (not nested.get("content_type") or nested.get("content_type") == "application/octet-stream"):
                nested["content_type"] = content_type
            return nested

        try:
            cleaned = re.sub(r"\s+", "", value_text)
            if (
                len(cleaned) >= 16
                and re.fullmatch(r"[A-Za-z0-9+/=]+", cleaned)
                and ("=" in cleaned[-2:] or len(cleaned) % 4 == 0)
            ):
                decoded = base64.b64decode(cleaned, validate=False)
                if decoded:
                    return {
                        "filename": filename,
                        "content": decoded,
                        "content_type": content_type,
                    }
        except (binascii.Error, ValueError):
            continue

    return None


def get_vendor_document_content_by_filename(file_name: str) -> dict[str, Any] | None:
    global conn

    normalized_name = Path(str(file_name or "")).name.strip()
    if not normalized_name:
        return None

    if not conn:
        initialize_database()
    if not conn:
        return None

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return None

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)
        col_map = {column.lower(): column for column in cols}

        filename_cols = []
        for column_name in cols:
            lowered = str(column_name).lower()
            if lowered.endswith("_filename") or "filename" in lowered:
                filename_cols.append(column_name)
        if not filename_cols:
            return None

        where_parts = [
            f"LOWER(TRIM(CAST({_qi(column_name)} AS VARCHAR(500)))) = LOWER(TRIM(%s))"
            for column_name in filename_cols
        ]
        query = f"SELECT * FROM {full_table} WHERE {' OR '.join(where_parts)} LIMIT 25"
        rows = cursor.fetchall()
        if not rows:
            return None

        for row in rows:
            row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}
            for filename_col in filename_cols:
                raw_filename = row_dict.get(filename_col)
                if Path(str(raw_filename or "")).name.strip().lower() != normalized_name.lower():
                    continue

                base_name = _document_column_base_name(filename_col)
                data_col = _resolve_column_name_loose(cols, f"{base_name}_Data")
                content_type_col = _resolve_column_name_loose(cols, f"{base_name}_ContentType")
                file_bytes = _coerce_binary_value(row_dict.get(data_col)) if data_col else None
                if file_bytes:
                    content_type = str(row_dict.get(content_type_col) or "application/octet-stream").strip()
                    if not content_type:
                        content_type = "application/octet-stream"
                    return {
                        "filename": normalized_name,
                        "content": file_bytes,
                        "content_type": content_type,
                    }

                text_payload = _extract_binary_from_text(row_dict.get(data_col) if data_col else row_dict.get(filename_col))
                if text_payload:
                    return {
                        "filename": Path(str(text_payload.get("filename") or normalized_name)).name.strip() or normalized_name,
                        "content": text_payload.get("content") or b"",
                        "content_type": str(text_payload.get("content_type") or "application/octet-stream"),
                    }
        return None
    except Exception:
        return None


def _build_document_column_candidates(column_name: str) -> list[str]:
    value = str(column_name or "").strip()
    if not value:
        return []

    base_name = _document_column_base_name(value).strip()
    candidates: list[str] = []
    for candidate in (
        value,
        f"{value}_Downloaded",
        f"{value}_Data",
        f"{value}_ContentType",
        f"{value}_FileName",
        base_name,
        f"{base_name}_Downloaded",
        f"{base_name}_Data",
        f"{base_name}_ContentType",
        f"{base_name}_FileName",
    ):
        normalized = str(candidate or "").strip()
        if normalized and normalized not in candidates:
            candidates.append(normalized)
    return candidates


def get_vendor_document_content_by_record_id(
    record_id: str,
    field_candidates: list[str],
) -> dict[str, Any] | None:
    global conn

    record_value = str(record_id or "").strip()
    if not record_value:
        return None

    requested_fields = [str(value or "").strip() for value in field_candidates if str(value or "").strip()]
    if not requested_fields:
        return None

    if not conn:
        initialize_database()
    if not conn:
        return None

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return None

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)
        col_map = {column.lower(): column for column in cols}
        record_col = _pick_col(cols, RECORD_COLS)
        if not record_col:
            return None

        query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({_qi(record_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s)) LIMIT 1"
        )
        cursor.execute(query, (record_value,))
        row = cursor.fetchone()
        if not row:
            return None

        row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}

        file_name_hints: list[str] = []
        content_type_hints: list[str] = []

        def append_filename_hint(raw_value: Any) -> None:
            value = Path(str(raw_value or "")).name.strip()
            if value and value not in file_name_hints:
                file_name_hints.append(value)

        def append_content_type_hint(raw_value: Any) -> None:
            value = str(raw_value or "").strip()
            if value and value not in content_type_hints:
                content_type_hints.append(value)

        def binary_payload_for_column(actual_column: str) -> dict[str, Any] | None:
            file_bytes = _coerce_binary_value(row_dict.get(actual_column))
            if not file_bytes:
                return None

            filename = ""
            for hint in file_name_hints:
                if hint:
                    filename = hint
                    break
            if not filename:
                filename = Path(actual_column).name.strip() or "document"

            content_type = ""
            for hint in content_type_hints:
                if hint:
                    content_type = hint
                    break
            if not content_type:
                content_type = "application/octet-stream"

            return {
                "filename": filename,
                "content": file_bytes,
                "content_type": content_type,
            }

        actual_document_columns: list[str] = []
        for field_name in requested_fields:
            for candidate in _build_document_column_candidates(field_name):
                actual_column = col_map.get(candidate.lower()) or _resolve_column_name_loose(cols, candidate)
                if actual_column and actual_column not in actual_document_columns:
                    actual_document_columns.append(actual_column)

        if not actual_document_columns:
            token_set: set[str] = set()
            for field_name in requested_fields:
                normalized = _normalize_column_key(field_name)
                if not normalized:
                    continue
                for token in re.findall(r"[a-z]+|\d+", normalized):
                    if len(token) >= 3:
                        token_set.add(token)

            doc_hints = ("file", "doc", "attach", "cert", "download", "data", "blob")
            for column_name in cols:
                lowered = str(column_name).lower()
                normalized_col = _normalize_column_key(column_name)
                if not any(hint in lowered for hint in doc_hints):
                    continue
                if token_set and not any(token in normalized_col for token in token_set):
                    continue
                actual_document_columns.append(column_name)

        for actual_column in actual_document_columns:
            raw_value = row_dict.get(actual_column)
            if isinstance(raw_value, str):
                if actual_column.lower().endswith("contenttype"):
                    append_content_type_hint(raw_value)
                else:
                    append_filename_hint(raw_value)

        for actual_column in actual_document_columns:
            payload = binary_payload_for_column(actual_column)
            if payload:
                return payload

        for actual_column in actual_document_columns:
            raw_value = row_dict.get(actual_column)
            if not isinstance(raw_value, str):
                continue
            text_payload = _extract_binary_from_text(raw_value)
            if not text_payload:
                continue
            filename = Path(str(text_payload.get("filename") or "")).name.strip()
            if not filename:
                for hint in file_name_hints:
                    if hint:
                        filename = hint
                        break
            if not filename:
                filename = Path(actual_column).name.strip() or "document"
            content_type = str(text_payload.get("content_type") or "").strip()
            if not content_type:
                content_type = "application/octet-stream"
            return {
                "filename": filename,
                "content": text_payload.get("content") or b"",
                "content_type": content_type,
            }

        return None
    except Exception:
        return None


def update_vendor_approver_status_by_record_id(record_id: str, approver_status: str) -> bool:
    global conn

    record_value = str(record_id or "").strip()
    status_value = str(approver_status or "").strip()
    if not record_value or not status_value:
        return False

    if not conn:
        initialize_database()
    if not conn:
        return False

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return False

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)

        record_col = _pick_col(cols, RECORD_COLS)
        status_col = _pick_col(cols, APPROVER_STATUS_COLS)
        updated_at_col = _pick_col(cols, ["updatedAt", "UpdatedAt"])
        if not record_col or not status_col:
            return False

        ist = timezone(timedelta(hours=5, minutes=30))
        now_str = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S")
        set_parts = [f"{_qi(status_col)} = %s"]
        params: list[Any] = [status_value]
        if updated_at_col:
            set_parts.append(f"{_qi(updated_at_col)} = %s")
            params.append(now_str)

        query = (
            f"UPDATE {full_table} "
            f"SET {', '.join(set_parts)} "
            f"WHERE LOWER(TRIM(CAST({_qi(record_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s))"
        )
        params.append(record_value)
        cursor.execute(query, params)
        affected_rows = cursor.rowcount
        conn.commit()
        return affected_rows > 0
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return False


def update_vendor_record_by_record_id(record_id: str, updates: dict[str, Any]) -> bool:
    global conn

    record_value = str(record_id or "").strip()
    if not record_value or not isinstance(updates, dict) or not updates:
        return False

    if not conn:
        initialize_database()
    if not conn:
        return False

    try:
        cursor = conn.cursor()
        vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
        if not vendor_table:
            return False

        schema, table = vendor_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)
        col_types = _list_column_types(cursor, schema, table)
        blocked_columns = _list_non_insertable_columns(cursor, schema, table)
        col_map = {c.lower(): c for c in cols}

        record_col = _pick_col(cols, RECORD_COLS)
        if not record_col:
            return False

        payload: dict[str, Any] = {}
        for key, value in updates.items():
            key_lower = str(key or "").strip().lower()
            if not key_lower:
                continue
            actual_col = col_map.get(key_lower)
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
            if actual_col.lower() in VALIDATION_FLAG_COLUMN_SET:
                normalized_flag = _normalize_validation_flag_value(value)
                if sql_type in {"bit", "boolean", "bool"}:
                    payload[actual_col] = None if normalized_flag is None else normalized_flag == "True"
                else:
                    payload[actual_col] = normalized_flag
            else:
                payload[actual_col] = _coerce_sql_value(value, sql_type)

        updated_at_col = _pick_col(cols, ["updatedAt", "UpdatedAt"])
        if (
            updated_at_col
            and updated_at_col.lower() not in blocked_columns
            and updated_at_col.lower() != record_col.lower()
        ):
            ist = timezone(timedelta(hours=5, minutes=30))
            payload[updated_at_col] = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S")

        if not payload:
            return False

        set_parts = [f"{_qi(column)} = %s" for column in payload.keys()]
        params = list(payload.values())
        params.append(record_value)

        query = (
            f"UPDATE {full_table} "
            f"SET {', '.join(set_parts)} "
            f"WHERE LOWER(TRIM(CAST({_qi(record_col)} AS VARCHAR(200)))) = LOWER(TRIM(%s))"
        )
        cursor.execute(query, params)
        affected_rows = cursor.rowcount
        conn.commit()
        return affected_rows > 0
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return False


def insert_db_user(
    name: str,
    email: str,
    password: str,
    role: str,
    bot_name: str | None = None,
) -> dict[str, Any]:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        raise RuntimeError(db_status.get("error") or "Database connection failed")

    try:
        cursor = conn.cursor()
        users_table = _find_table(cursor, ["Users"])
        if not users_table:
            raise RuntimeError("Users table not found")

        schema, table = users_table
        full_table = _qt(schema, table)
        cols = _list_columns(cursor, schema, table)
        col_types = _list_column_types(cursor, schema, table)
        blocked_columns = _list_non_insertable_columns(cursor, schema, table)
        col_map = {c.lower(): c for c in cols}

        email_col = _pick_col(cols, EMAIL_COLS)
        pass_col = _pick_col(cols, PASSWORD_COLS)
        name_col = _pick_col(cols, NAME_COLS)
        role_col = _pick_col(cols, ROLE_COLS)
        bot_col = _pick_col(cols, BOTNAME_COLS)
        id_col = _pick_col(cols, ID_COLS)

        if not email_col or not pass_col:
            raise RuntimeError("Required columns not found in Users table (email/password)")
        if email_col.lower() in blocked_columns or pass_col.lower() in blocked_columns:
            raise RuntimeError("Users table email/password columns are not insertable")

        duplicate_query = (
            f"SELECT 1 FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({_qi(email_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s)) LIMIT 1"
        )
        cursor.execute(duplicate_query, (email,))
        if cursor.fetchone():
            raise ValueError("User with this email already exists")

        clean_name = str(name or "").strip()
        clean_email = str(email or "").strip()
        clean_role = str(role or "").strip()
        password_value = _hash_password_for_storage(password)
        payload: dict[str, Any] = {
            email_col: clean_email,
            pass_col: password_value,
        }

        if name_col and name_col.lower() not in blocked_columns:
            payload[name_col] = clean_name
        if role_col and role_col.lower() not in blocked_columns:
            payload[role_col] = clean_role
        if bot_col and bot_col.lower() not in blocked_columns:
            payload[bot_col] = str(bot_name or "Vendor Bot").strip()

        ist = timezone(timedelta(hours=5, minutes=30))
        now_str = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S")
        for time_col_key in ("createdat", "updatedat"):
            actual_col = col_map.get(time_col_key)
            if actual_col and actual_col.lower() not in blocked_columns and actual_col not in payload:
                sql_type = col_types.get(actual_col.lower(), "")
                payload[actual_col] = _coerce_sql_value(now_str, sql_type) or now_str

        required_columns = _list_required_non_nullable_columns(cursor, schema, table)
        for column_meta in required_columns:
            column_name = str(column_meta.get("name") or "").strip()
            if not column_name:
                continue
            if column_name.lower() in blocked_columns:
                continue
            if column_name in payload:
                continue

            sql_type = str(column_meta.get("data_type") or "").strip().lower()
            fallback_value = _fallback_required_user_column_value(
                column_name,
                sql_type,
                name=clean_name,
                email=clean_email,
                role=clean_role,
                bot_name=bot_name,
                password_value=password_value,
                now_value=now_str,
            )
            payload[column_name] = _coerce_sql_value(fallback_value, sql_type)

        cols_sql = ", ".join(f"{_qi(c)}" for c in payload.keys())
        params_sql = ", ".join("%s" for _ in payload.values())
        values = list(payload.values())
        cursor.execute(f"INSERT INTO {full_table} ({cols_sql}) VALUES ({params_sql})", values)
        conn.commit()

        select_query = (
            f"SELECT * FROM {full_table} "
            f"WHERE LOWER(TRIM(CAST({_qi(email_col)} AS VARCHAR(320)))) = LOWER(TRIM(%s))"
        )
        if id_col:
            select_query += f" ORDER BY {_qi(id_col)} DESC"
        select_query += " LIMIT 1"
        cursor.execute(select_query, (email,))
        row = cursor.fetchone()
        if not row:
            return {"id": None, "name": name, "email": email, "role": role}

        row_dict = {desc[0]: row[idx] for idx, desc in enumerate(cursor.description)}
        mapped = _map_user_row(row_dict, id_col, name_col, email_col, role_col, bot_col=bot_col)
        if mapped:
            return mapped
        return {"id": None, "name": name, "email": email, "role": role}
    except ValueError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Unable to create user in SQL Users table: {exc}") from exc


def insert_vendor_master_data(vendor_data: dict[str, Any]) -> str:
    global conn

    if not conn:
        initialize_database()
    if not conn:
        raise RuntimeError(db_status.get("error") or "Database connection failed")

    cursor = conn.cursor()

    vendor_table = _find_table(cursor, VENDOR_TABLE_NAMES)
    if not vendor_table:
        raise RuntimeError("Vendor master table not found")

    schema, table = vendor_table
    full_table = _qt(schema, table)

    columns = _list_columns(cursor, schema, table)
    col_types = _list_column_types(cursor, schema, table)
    blocked_columns = _list_non_insertable_columns(cursor, schema, table)
    col_map = {c.lower(): c for c in columns}

    payload: dict[str, Any] = {}
    for key, value in vendor_data.items():
        if value is None:
            continue

        # psycopg2 can only bind scalar values; convert list/dict payloads to text.
        if isinstance(value, (list, dict, tuple, set)):
            value = str(value)

        key_lower = key.lower()
        if key_lower in col_map:
            actual_col = col_map[key_lower]
            if actual_col.lower() in blocked_columns:
                continue
            sql_type = col_types.get(actual_col.lower(), "")
            if actual_col.lower() in VALIDATION_FLAG_COLUMN_SET:
                normalized_flag = _normalize_validation_flag_value(value)
                if sql_type in {"bit", "boolean", "bool"}:
                    payload[actual_col] = None if normalized_flag is None else normalized_flag == "True"
                else:
                    payload[actual_col] = normalized_flag
            else:
                payload[actual_col] = _coerce_sql_value(value, sql_type)

    ist = timezone(timedelta(hours=5, minutes=30))
    now_str = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S")
    if "createdat" in col_map and col_map["createdat"] not in payload:
        payload[col_map["createdat"]] = now_str
    if "updatedat" in col_map and col_map["updatedat"] not in payload:
        payload[col_map["updatedat"]] = now_str

    record_id_col = col_map.get("recordid")
    if record_id_col and record_id_col.lower() in blocked_columns:
        record_id_col = None
    record_id_type = col_types.get("recordid", "")
    requested_record_id = ""
    if record_id_col and record_id_col in payload:
        requested_record_id = str(payload.get(record_id_col) or "").strip()
    generated_record_id = None
    last_error: Exception | None = None

    for attempt in range(6):
        local_payload = dict(payload)
        if record_id_col:
            if requested_record_id and attempt == 0:
                generated_record_id = requested_record_id
            elif record_id_type in {"uniqueidentifier", "uuid"}:
                generated_record_id = str(uuid.uuid4())
            else:
                generated_record_id = _generate_record_id(24)
            local_payload[record_id_col] = generated_record_id

        if not local_payload:
            raise RuntimeError("No matching columns found to insert")

        cols_sql = ", ".join(f"{_qi(c)}" for c in local_payload.keys())
        params_sql = ", ".join("%s" for _ in local_payload.values())
        values = list(local_payload.values())

        try:
            cursor.execute(f"INSERT INTO {full_table} ({cols_sql}) VALUES ({params_sql})", values)
            conn.commit()
            if generated_record_id:
                return generated_record_id
            return ""
        except IntegrityError as exc:
            last_error = exc
            message = str(exc).lower()
            # Retry only for duplicate key collisions where a new RecordID could succeed.
            if "duplicate" in message or "unique" in message or "2627" in message or "2601" in message:
                continue
            raise RuntimeError(f"SQL integrity error: {exc}") from exc
        except Exception as exc:
            raise RuntimeError(f"SQL insert error: {exc}") from exc

    if last_error:
        raise RuntimeError(f"Unable to insert vendor data after retries: {last_error}") from last_error
    raise RuntimeError("Unable to insert vendor data after retries")
