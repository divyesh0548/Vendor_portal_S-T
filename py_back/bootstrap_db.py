"""
Create the RDS PostgreSQL database (if missing) and required portal tables.

Usage (from py_back, with app/.env configured):
    python bootstrap_db.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import psycopg2
from dotenv import load_dotenv
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

APP_DIR = Path(__file__).resolve().parent / "app"
load_dotenv(APP_DIR / ".env")
load_dotenv(Path(__file__).resolve().parent / ".env")


def _env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def _connect_kwargs(dbname: str) -> dict:
    host = _env("RDS_HOST")
    user = _env("RDS_USER")
    password = os.getenv("RDS_PASSWORD") or ""
    port = _env("RDS_PORT", "5432")
    sslmode = _env("RDS_SSLMODE", "require") or "require"
    if not host or not user:
        raise SystemExit("Missing RDS_HOST / RDS_USER in .env")
    return {
        "host": host,
        "port": int(port) if port.isdigit() else port,
        "user": user,
        "password": password,
        "dbname": dbname,
        "sslmode": sslmode,
        "connect_timeout": 15,
    }


USERS_DDL = """
CREATE TABLE IF NOT EXISTS "Users" (
    "Id" BIGSERIAL PRIMARY KEY,
    "UserName" VARCHAR(100) NOT NULL,
    "Email" VARCHAR(150) NOT NULL UNIQUE,
    "BOTName" VARCHAR(100) NULL,
    "Role" VARCHAR(50) NOT NULL,
    "Password" VARCHAR(255) NOT NULL
);
"""

VENDOR_DDL = """
CREATE TABLE IF NOT EXISTS "VendorMasterData" (
    "Id" BIGSERIAL,
    "recordId" VARCHAR(64) PRIMARY KEY,
    "companyName" VARCHAR(500) NULL,
    "vendorName" VARCHAR(500) NULL,
    "vendorCategory" VARCHAR(500) NULL,
    "gstRegistrationStatus" VARCHAR(500) NULL,
    "gstNumber" VARCHAR(50) NULL,
    "tradeNameGst" VARCHAR(500) NULL,
    "gstDealerType" VARCHAR(100) NULL,
    "gstReturnFiling" VARCHAR(50) NULL,
    "invoiceApplicable" VARCHAR(10) NULL,
    "eInvoiceApplicability" VARCHAR(10) NULL,
    "eWayBillApplicability" VARCHAR(10) NULL,
    "gstDocument" BYTEA NULL,
    "gstDocument_FileName" VARCHAR(512) NULL,
    "gstDocument_ContentType" VARCHAR(255) NULL,
    "gstDocument_FileSize" BIGINT NULL,
    "addressLane1" VARCHAR(500) NULL,
    "addressLane2" VARCHAR(500) NULL,
    "addressLane3" VARCHAR(500) NULL,
    "pincode" VARCHAR(20) NULL,
    "country" VARCHAR(500) NULL,
    "state" VARCHAR(500) NULL,
    "district" VARCHAR(500) NULL,
    "company" VARCHAR(500) NULL,
    "cinNumber" VARCHAR(50) NULL,
    "cinDocument" BYTEA NULL,
    "cinDocument_FileName" VARCHAR(512) NULL,
    "cinDocument_ContentType" VARCHAR(255) NULL,
    "cinDocument_FileSize" BIGINT NULL,
    "panStatus" VARCHAR(500) NULL,
    "panNumber" VARCHAR(50) NULL,
    "businessEntityType" VARCHAR(500) NULL,
    "tradeNamePan" VARCHAR(500) NULL,
    "itrFilingStatus" VARCHAR(10) NULL,
    "panDocument" BYTEA NULL,
    "panDocument_FileName" VARCHAR(512) NULL,
    "panDocument_ContentType" VARCHAR(255) NULL,
    "panDocument_FileSize" BIGINT NULL,
    "msmeStatus" VARCHAR(500) NULL,
    "udyamNumber" VARCHAR(500) NULL,
    "msmeCategory" VARCHAR(500) NULL,
    "msmeIndustry" VARCHAR(500) NULL,
    "msmeDocument" BYTEA NULL,
    "msmeDocument_FileName" VARCHAR(512) NULL,
    "msmeDocument_ContentType" VARCHAR(255) NULL,
    "msmeDocument_FileSize" BIGINT NULL,
    "pfStatus" VARCHAR(500) NULL,
    "pfNumber" VARCHAR(500) NULL,
    "pfDocument" BYTEA NULL,
    "pfDocument_FileName" VARCHAR(512) NULL,
    "pfDocument_ContentType" VARCHAR(255) NULL,
    "pfDocument_FileSize" BIGINT NULL,
    "esiStatus" VARCHAR(500) NULL,
    "esiNumber" VARCHAR(500) NULL,
    "esiDocument" BYTEA NULL,
    "esiDocument_FileName" VARCHAR(512) NULL,
    "esiDocument_ContentType" VARCHAR(255) NULL,
    "esiDocument_FileSize" BIGINT NULL,
    "department" VARCHAR(500) NULL,
    "buyerId" VARCHAR(500) NULL,
    "bankPaymentMethod" VARCHAR(500) NULL,
    "ifscCode" VARCHAR(500) NULL,
    "bankName" VARCHAR(500) NULL,
    "branchName" VARCHAR(500) NULL,
    "bankAccountNumber" VARCHAR(500) NULL,
    "cancelledCheque" BYTEA NULL,
    "cancelledCheque_FileName" VARCHAR(512) NULL,
    "cancelledCheque_ContentType" VARCHAR(255) NULL,
    "cancelledCheque_FileSize" BIGINT NULL,
    "statutoryDetails" TEXT NULL,
    "totalEmployees" VARCHAR(500) NULL,
    "inhouseMachineries" VARCHAR(500) NULL,
    "inhouseMachineriesDoc" BYTEA NULL,
    "inhouseMachineriesDoc_FileName" VARCHAR(512) NULL,
    "inhouseMachineriesDoc_ContentType" VARCHAR(255) NULL,
    "inhouseMachineriesDoc_FileSize" BIGINT NULL,
    "numberOfPlants" VARCHAR(500) NULL,
    "numberOfPlantsDoc" BYTEA NULL,
    "numberOfPlantsDoc_FileName" VARCHAR(512) NULL,
    "numberOfPlantsDoc_ContentType" VARCHAR(255) NULL,
    "numberOfPlantsDoc_FileSize" BIGINT NULL,
    "yearOfIncorporation" VARCHAR(500) NULL,
    "selectedPreviousYear" VARCHAR(500) NULL,
    "previousYearTurnover" VARCHAR(100) NULL,
    "headOfficeLocation" VARCHAR(500) NULL,
    "materialDealingsIn" TEXT NULL,
    "serviceProvideFor" TEXT NULL,
    "isoCertificates" VARCHAR(500) NULL,
    "iso1" BYTEA NULL,
    "iso1_FileName" VARCHAR(512) NULL,
    "iso1_ContentType" VARCHAR(255) NULL,
    "iso1_FileSize" BIGINT NULL,
    "iso2" BYTEA NULL,
    "iso2_FileName" VARCHAR(512) NULL,
    "iso2_ContentType" VARCHAR(255) NULL,
    "iso2_FileSize" BIGINT NULL,
    "iso3" BYTEA NULL,
    "iso3_FileName" VARCHAR(512) NULL,
    "iso3_ContentType" VARCHAR(255) NULL,
    "iso3_FileSize" BIGINT NULL,
    "iso4" BYTEA NULL,
    "iso4_FileName" VARCHAR(512) NULL,
    "iso4_ContentType" VARCHAR(255) NULL,
    "iso4_FileSize" BIGINT NULL,
    "iso5" BYTEA NULL,
    "iso5_FileName" VARCHAR(512) NULL,
    "iso5_ContentType" VARCHAR(255) NULL,
    "iso5_FileSize" BIGINT NULL,
    "ecoVadisScore" VARCHAR(500) NULL,
    "ecoVadisDoc" BYTEA NULL,
    "ecoVadisDoc_FileName" VARCHAR(512) NULL,
    "ecoVadisDoc_ContentType" VARCHAR(255) NULL,
    "ecoVadisDoc_FileSize" BIGINT NULL,
    "ohsasCertificate" VARCHAR(500) NULL,
    "ohsasDoc" BYTEA NULL,
    "ohsasDoc_FileName" VARCHAR(512) NULL,
    "ohsasDoc_ContentType" VARCHAR(255) NULL,
    "ohsasDoc_FileSize" BIGINT NULL,
    "relevantCertificate" VARCHAR(500) NULL,
    "relevantCertificateDoc" BYTEA NULL,
    "relevantCertificateDoc_FileName" VARCHAR(512) NULL,
    "relevantCertificateDoc_ContentType" VARCHAR(255) NULL,
    "relevantCertificateDoc_FileSize" BIGINT NULL,
    "contactPersonName" VARCHAR(500) NULL,
    "contactPersonDesignation" VARCHAR(500) NULL,
    "contactPersonEmail" VARCHAR(320) NULL,
    "contactPersonMobile" VARCHAR(500) NULL,
    "alternativePersonName" VARCHAR(500) NULL,
    "alternativePersonDesignation" VARCHAR(500) NULL,
    "alternativePersonEmail" VARCHAR(320) NULL,
    "alternativePersonMobile" VARCHAR(500) NULL,
    "alternativePersonDate" DATE NULL,
    "alternativePersonPlace" VARCHAR(500) NULL,
    "agreeCodeOfConduct" VARCHAR(20) NULL,
    "agreeInformationAccuracy" VARCHAR(20) NULL,
    "q1" VARCHAR(10) NULL,
    "q2" VARCHAR(10) NULL,
    "q3" VARCHAR(10) NULL,
    "q4" VARCHAR(10) NULL,
    "q5" VARCHAR(10) NULL,
    "q6" VARCHAR(10) NULL,
    "q7" VARCHAR(10) NULL,
    "q8" VARCHAR(10) NULL,
    "q9" VARCHAR(10) NULL,
    "incoTerms" VARCHAR(500) NULL,
    "paymentTerms" VARCHAR(500) NULL,
    "vendorType" VARCHAR(500) NULL,
    "rubaminApproverHod" VARCHAR(500) NULL,
    "currency" VARCHAR(500) NULL,
    "salesContract" VARCHAR(500) NULL,
    "form10fStatus" VARCHAR(500) NULL,
    "form10fDocument" BYTEA NULL,
    "form10fDocument_FileName" VARCHAR(512) NULL,
    "form10fDocument_ContentType" VARCHAR(255) NULL,
    "form10fDocument_FileSize" BIGINT NULL,
    "trcDocument" BYTEA NULL,
    "trcDocument_FileName" VARCHAR(512) NULL,
    "trcDocument_ContentType" VARCHAR(255) NULL,
    "trcDocument_FileSize" BIGINT NULL,
    "importDocument1" BYTEA NULL,
    "importDocument1_FileName" VARCHAR(512) NULL,
    "importDocument1_ContentType" VARCHAR(255) NULL,
    "importDocument1_FileSize" BIGINT NULL,
    "importDocument2" BYTEA NULL,
    "importDocument2_FileName" VARCHAR(512) NULL,
    "importDocument2_ContentType" VARCHAR(255) NULL,
    "importDocument2_FileSize" BIGINT NULL,
    "importDocument3" BYTEA NULL,
    "importDocument3_FileName" VARCHAR(512) NULL,
    "importDocument3_ContentType" VARCHAR(255) NULL,
    "importDocument3_FileSize" BIGINT NULL,
    "beneficiaryName" VARCHAR(500) NULL,
    "swiftCode" VARCHAR(500) NULL,
    "iban" VARCHAR(500) NULL,
    "bankDocument" BYTEA NULL,
    "bankDocument_FileName" VARCHAR(512) NULL,
    "bankDocument_ContentType" VARCHAR(255) NULL,
    "bankDocument_FileSize" BIGINT NULL,
    "VendorType" VARCHAR(50) NULL,
    "IncoTerms" VARCHAR(200) NULL,
    "PaymentTerms" VARCHAR(500) NULL,
    "WithholdingTax" TEXT NULL,
    "TurnoverLimit" VARCHAR(50) NULL,
    "RubaminApproverHod" VARCHAR(320) NULL,
    "Rubamin Contact Person" VARCHAR(500) NULL,
    "Rubamin Contact Person Email" VARCHAR(320) NULL,
    "Rubamin Contact Person Department" VARCHAR(200) NULL,
    "ApproverStatus" VARCHAR(50) NULL,
    "approverStatus" VARCHAR(50) NULL,
    "ReviewDecision" VARCHAR(50) NULL,
    "ReviewRemarks" TEXT NULL,
    "ReviewedByRole" VARCHAR(100) NULL,
    "VendorCode" VARCHAR(50) NULL,
    "SAPCodeGeneratedAt" VARCHAR(50) NULL,
    "TaxpayerTypeGST" VARCHAR(200) NULL,
    "GST Validate" VARCHAR(5) NULL,
    "CIN Validate" VARCHAR(5) NULL,
    "PAN Validate" VARCHAR(5) NULL,
    "MSME Validate" VARCHAR(5) NULL,
    "PF Validate" VARCHAR(5) NULL,
    "ESI Validate" VARCHAR(5) NULL,
    "Address Validate" VARCHAR(5) NULL,
    "Bank Validate" VARCHAR(5) NULL,
    "Validate By" VARCHAR(320) NULL,
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    "updated_at" TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    "createdAt" TIMESTAMPTZ NULL DEFAULT NOW(),
    "updatedAt" TIMESTAMPTZ NULL DEFAULT NOW()
);
"""

CUSTOMER_DDL = """
CREATE TABLE IF NOT EXISTS "CustomerMasterData" (
    "recordId" VARCHAR(64) PRIMARY KEY,
    "CompanyDealing" VARCHAR(500) NULL,
    "CustomerName" VARCHAR(500) NULL,
    "CompanyStatus" VARCHAR(500) NULL,
    "FirmType" VARCHAR(500) NULL,
    "ManufacturingType" VARCHAR(500) NULL,
    "TypeofEntity" VARCHAR(500) NULL,
    "PANStatus" VARCHAR(500) NULL,
    "PANNumber" VARCHAR(20) NULL,
    "TradeNamePAN" VARCHAR(500) NULL,
    "TAN" VARCHAR(500) NULL,
    "CIN" VARCHAR(30) NULL,
    "PANUpload" BYTEA NULL,
    "PANUpload_FileName" VARCHAR(512) NULL,
    "PANUpload_ContentType" VARCHAR(255) NULL,
    "PANUpload_FileSize" BIGINT NULL,
    "IncorporationCertificate" BYTEA NULL,
    "IncorporationCertificate_FileName" VARCHAR(512) NULL,
    "IncorporationCertificate_ContentType" VARCHAR(255) NULL,
    "IncorporationCertificate_FileSize" BIGINT NULL,
    "GSTStatus" VARCHAR(500) NULL,
    "GSTIN" VARCHAR(20) NULL,
    "TradeNameGST" VARCHAR(500) NULL,
    "GSTCertificate" BYTEA NULL,
    "GSTCertificate_FileName" VARCHAR(512) NULL,
    "GSTCertificate_ContentType" VARCHAR(255) NULL,
    "GSTCertificate_FileSize" BIGINT NULL,
    "GSTIN1" VARCHAR(20) NULL,
    "TradeNameGST1" VARCHAR(500) NULL,
    "GSTCertificate1" BYTEA NULL,
    "GSTCertificate1_FileName" VARCHAR(512) NULL,
    "GSTCertificate1_ContentType" VARCHAR(255) NULL,
    "GSTCertificate1_FileSize" BIGINT NULL,
    "GSTIN2" VARCHAR(20) NULL,
    "TradeNameGST2" VARCHAR(500) NULL,
    "GSTCertificate2" BYTEA NULL,
    "GSTCertificate2_FileName" VARCHAR(512) NULL,
    "GSTCertificate2_ContentType" VARCHAR(255) NULL,
    "GSTCertificate2_FileSize" BIGINT NULL,
    "MaterialDealing" TEXT NULL,
    "ServiceDealing" TEXT NULL,
    "Country" VARCHAR(200) NULL,
    "State" VARCHAR(200) NULL,
    "State_Name" VARCHAR(200) NULL,
    "District" VARCHAR(200) NULL,
    "Pincode" VARCHAR(20) NULL,
    "Shipto" VARCHAR(20) NULL,
    "Billto" VARCHAR(20) NULL,
    "State1" VARCHAR(200) NULL,
    "StateName1" VARCHAR(200) NULL,
    "State2" VARCHAR(200) NULL,
    "Lane1" VARCHAR(500) NULL,
    "Lane2" VARCHAR(500) NULL,
    "Lane3" VARCHAR(500) NULL,
    "Lane4" VARCHAR(500) NULL,
    "Lane5" VARCHAR(500) NULL,
    "Lane6" VARCHAR(500) NULL,
    "Lane7" VARCHAR(500) NULL,
    "Lane8" VARCHAR(500) NULL,
    "Lane9" VARCHAR(500) NULL,
    "District1" VARCHAR(200) NULL,
    "District2" VARCHAR(200) NULL,
    "Pincode1" VARCHAR(20) NULL,
    "Pincode2" VARCHAR(20) NULL,
    "FurnishBy" VARCHAR(500) NULL,
    "Designation" VARCHAR(500) NULL,
    "Email" VARCHAR(320) NULL,
    "Mobile" VARCHAR(30) NULL,
    "Datepicker" DATE NULL,
    "Place" VARCHAR(500) NULL,
    "AP_FurnishBy1" VARCHAR(500) NULL,
    "AP_Designation1" VARCHAR(500) NULL,
    "AP_Email1" VARCHAR(320) NULL,
    "AP_Mobile1" VARCHAR(30) NULL,
    "IncoTerms" VARCHAR(500) NULL,
    "PaymentTerms" VARCHAR(500) NULL,
    "CustomerGroup" VARCHAR(500) NULL,
    "HODEmail" VARCHAR(320) NULL,
    "hodEmail" VARCHAR(320) NULL,
    "SalesOrganization" VARCHAR(500) NULL,
    "DistributionChannel" VARCHAR(500) NULL,
    "Division" VARCHAR(500) NULL,
    "AccountAssignment" VARCHAR(500) NULL,
    "IndustryType" VARCHAR(500) NULL,
    "ScrapSales" VARCHAR(500) NULL,
    "CustomerType" VARCHAR(500) NULL,
    "Currency" VARCHAR(500) NULL,
    "InsuranceGroup" VARCHAR(500) NULL,
    "InsuranceLimit" DECIMAL(18,2) NULL,
    "CreditLimit" DECIMAL(18,2) NULL,
    "FileUploader" BYTEA NULL,
    "FileUploader_FileName" VARCHAR(512) NULL,
    "FileUploader_ContentType" VARCHAR(255) NULL,
    "FileUploader_FileSize" BIGINT NULL,
    "FileUploader1" BYTEA NULL,
    "FileUploader1_FileName" VARCHAR(512) NULL,
    "FileUploader1_ContentType" VARCHAR(255) NULL,
    "FileUploader1_FileSize" BIGINT NULL,
    "FileUploader2" BYTEA NULL,
    "FileUploader2_FileName" VARCHAR(512) NULL,
    "FileUploader2_ContentType" VARCHAR(255) NULL,
    "FileUploader2_FileSize" BIGINT NULL,
    "FileUploader3" BYTEA NULL,
    "FileUploader3_FileName" VARCHAR(512) NULL,
    "FileUploader3_ContentType" VARCHAR(255) NULL,
    "FileUploader3_FileSize" BIGINT NULL,
    "FileUploader4" BYTEA NULL,
    "FileUploader4_FileName" VARCHAR(512) NULL,
    "FileUploader4_ContentType" VARCHAR(255) NULL,
    "FileUploader4_FileSize" BIGINT NULL,
    "FileUploader5" BYTEA NULL,
    "FileUploader5_FileName" VARCHAR(512) NULL,
    "FileUploader5_ContentType" VARCHAR(255) NULL,
    "FileUploader5_FileSize" BIGINT NULL,
    "approverStatus" VARCHAR(200) NULL,
    "customerEmail" VARCHAR(320) NULL,
    "rubaminContactPersonEmail" VARCHAR(320) NULL,
    "rubaminContactPerson" VARCHAR(320) NULL,
    "rubaminContactPersonDepartment" VARCHAR(320) NULL,
    "reviewRemarks" TEXT NULL,
    "reviewedByRole" VARCHAR(100) NULL,
    "remarks" TEXT NULL,
    "withholdingTax" TEXT NULL,
    "tcsRate" VARCHAR(200) NULL,
    "turnoverLimit" VARCHAR(200) NULL,
    "gstValidate" VARCHAR(5) NULL,
    "GST Valid" VARCHAR(30) NULL,
    "TaxpayerTypeGST" VARCHAR(200) NULL,
    "cinValidate" VARCHAR(5) NULL,
    "panValidate" VARCHAR(5) NULL,
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    "updated_at" TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
"""

SEED_ADMIN_SQL = """
INSERT INTO "Users" ("UserName", "Email", "BOTName", "Role", "Password")
SELECT %s, %s, %s, %s, %s
WHERE NOT EXISTS (
    SELECT 1 FROM "Users"
    WHERE LOWER(TRIM("Email")) = LOWER(TRIM(%s))
);
"""


def ensure_database(target_db: str) -> None:
    admin_db = _env("RDS_ADMIN_DB", "postgres") or "postgres"
    print(f"Connecting to maintenance DB '{admin_db}' on {_env('RDS_HOST')}...")
    conn = psycopg2.connect(**_connect_kwargs(admin_db))
    conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM pg_database WHERE datname = %s", (target_db,))
            exists = cur.fetchone() is not None
            if exists:
                print(f"Database '{target_db}' already exists.")
            else:
                cur.execute(f'CREATE DATABASE "{target_db.replace(chr(34), "")}"')
                print(f"Created database '{target_db}'.")
    finally:
        conn.close()


def ensure_tables(target_db: str) -> None:
    print(f"Connecting to '{target_db}' and creating tables...")
    conn = psycopg2.connect(**_connect_kwargs(target_db))
    try:
        with conn.cursor() as cur:
            cur.execute(USERS_DDL)
            cur.execute(VENDOR_DDL)
            cur.execute(CUSTOMER_DDL)

            seed_email = _env("BOOTSTRAP_ADMIN_EMAIL", "admin@example.com")
            seed_name = _env("BOOTSTRAP_ADMIN_NAME", "Admin")
            seed_role = _env("BOOTSTRAP_ADMIN_ROLE", "Admin")
            seed_bot = _env("BOOTSTRAP_ADMIN_BOT", "PortalBot")
            seed_password = _env("BOOTSTRAP_ADMIN_PASSWORD", "ChangeMe@123")
            cur.execute(
                SEED_ADMIN_SQL,
                (seed_name, seed_email, seed_bot, seed_role, seed_password, seed_email),
            )
        conn.commit()
        print("Tables ready: Users, VendorMasterData, CustomerMasterData")
        print(f"Admin user ensured: {seed_email}")
    finally:
        conn.close()


def main() -> int:
    target_db = _env("RDS_DB_NAME", "vendor_portal")
    if not target_db:
        print("RDS_DB_NAME is required", file=sys.stderr)
        return 1
    try:
        ensure_database(target_db)
        ensure_tables(target_db)
    except Exception as exc:
        print(f"Bootstrap failed: {exc}", file=sys.stderr)
        return 1
    print("Bootstrap completed successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
