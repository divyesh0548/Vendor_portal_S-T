# Python Backend

FastAPI backend for the vendor and customer portal. It serves the browser pages in `app/templates`, exposes the portal API routes, stores records in SQL Server through `pyodbc`, stores uploaded documents under `app/uploaded_files`, and includes helper flows for Excel/SAP automation plus GST/MSME validation.

## Requirements

- Python 3.10 or newer
- PostgreSQL (local or AWS RDS) reachable from the app host
- Microsoft Edge or Google Chrome for HTML-to-PDF export
- Windows, Microsoft Excel, Microsoft Word, and SAP GUI for the SAP/vendor-upload automation paths
- Google Chrome for the Selenium validation scripts
- Tesseract OCR installed separately when running the MSME validation script

The Python packages are listed in `requirements.txt`. OCR and ML packages such as `torch`, `transformers`, and `easyocr` are included because the validation scripts import them directly.

## Run

```powershell
cd py_back
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 3000
```

Open the portal at:

- `http://localhost:3000/web/login`
- `http://localhost:3000/web/select-role`
- `http://localhost:3000/web/vendor-dashboard`
- `http://localhost:3000/web/customer-dashboard`

## Configuration

### Core

- `JWT_SECRET`: secret used to sign login tokens. Default is a development value.
- `SEED_SAMPLE_USERS`: enables fallback sample users when SQL auth is unavailable. Default is `true`.
- `PUBLIC_BASE_URL`: optional public URL used when generating links in email invitations.

### Database

The backend uses PostgreSQL via `psycopg2`. Configure `py_back/app/.env`:

```
RDS_HOST=your-rds-endpoint.amazonaws.com
RDS_PORT=5432
RDS_USER=postgres
RDS_PASSWORD=your-password
RDS_DB_NAME=vendor_portal
RDS_SSLMODE=require
```

Create the database and tables (idempotent):

```powershell
cd py_back
pip install -r requirements.txt
python bootstrap_db.py
```

Useful health check:

- `GET /api/db/health`

### SMTP

Used for vendor/customer invitation and approval emails.

- `SMTP_SENDER_EMAIL`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`
- `SMTP_HOST`, default `smtp.gmail.com`
- `SMTP_PORT`, default `587`
- `SMTP_USE_STARTTLS`, default `true`
- `SMTP_TIMEOUT_SECONDS`, default `30`

Fallback password variable names are also accepted: `GMAIL_APP_PASSWORD`, `GOOGLE_APP_PASSWORD`, and `MAIL_APP_PASSWORD`. Legacy `OUTLOOK_*` variables are still accepted by the code.

For Gmail, enable 2-Step Verification, generate an app password, and set it in `SMTP_PASSWORD`.

### SAP And Excel

Vendor SAP upload uses the workbooks in `Vendor Template` plus an external upload workbook.

- `SAP_VENDOR_UPLOAD_PATH`, default `D:\Vendor Upload Template - BOT.xlsx`
- `SAP_OPEN_VENDOR_UPLOAD_AFTER_SYNC`, default `false`
- `SAP_CLOSE_VENDOR_UPLOAD_AFTER_SYNC`, default `true`
- `SAP_LOGON_EXECUTABLE_PATH`, default `C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe`
- `SAP_AUTO_LAUNCH_LOGON`, default `true`
- `SAP_AUTO_LOGON_AFTER_LAUNCH`, default `true`
- `SAP_SERVER_NAME`, default `DEV Server`
- `SAP_LOGON_USERNAME`, default `SAPAudit5`
- `SAP_LOGON_PASSWORD`: set this locally instead of relying on development defaults
- `SAP_LOGON_CLIENT`, default `400`
- `SAP_LOGON_LANGUAGE`, default `EN`
- `SAP_VENDOR_UPLOAD_TRANSACTION_CODE`, default `ZVEND_UP`
- `SAP_GUI_WAIT_TIMEOUT_SECONDS`, default `120`
- `SAP_TRANSACTION_WAIT_SECONDS`, default `20`
- `SAP_EXCEL_MINIMIZE_DELAY_SECONDS`, default `3`

SAP automation requires SAP GUI scripting to be available in the desktop session where the server is running.

## Main Pages

- `/web/login`
- `/web/change-password`
- `/web/select-role`
- `/web/admin-users`
- `/web/vendor-dashboard`
- `/web/invite-vendor`
- `/web/vendor-registration`
- `/web/vendor-registration/prescreen`
- `/web/vendor-registration/details`
- `/web/internal-registration`
- `/web/import-vendor`
- `/web/vendor-history`
- `/web/view-vendor`
- `/web/update-vendor`
- `/web/customer-dashboard`
- `/web/invite-customer`
- `/web/customer-public-form`
- `/web/customer-form`
- `/web/customer-shipto`
- `/web/view-customer`
- `/web/update-customer`

## API Groups

- `/api/auth`
- `/api/admin`
- `/api/vendor`
- `/api/customer`
- `/api/buyer`
- `/api/banking`
- `/api/accounts`
- `/api/hod`
- `/api/taxation`
- `/api/approval-sequence`
- `/api/db/health`
- `/api/pincode/{pincode}`
- `/api/ifsc/{ifsc}`

## Files And Documents

Uploaded files are saved below `app/uploaded_files` and are served through `/mock-files/{path}` or through the vendor/customer document endpoints. The `/mock-files` URL name is kept for compatibility with the older frontend route naming.

Generated exports can include Excel, PDF, and document ZIP downloads depending on the endpoint.

## Validation Scripts

GST validation service:

```powershell
cd "Data Validation\GST Extraction Process"
python -m uvicorn api:app --host 0.0.0.0 --port 8000 --reload
```

MSME/Udyam validation script:

```powershell
python "Data Validation\MSME Validation\Udyam Data Extract.py"
```

If Tesseract is installed in a different location, update `pytesseract.pytesseract.tesseract_cmd` in the MSME script.

## Troubleshooting

- If `pyodbc` connects fail, confirm the SQL Server ODBC driver is installed and `SQLSERVER_CONNECTION_STRING` points to the correct server/database.
- If PDF export fails, install Microsoft Edge or Google Chrome.
- If SAP upload fails, confirm SAP GUI is installed, SAP scripting is enabled, Excel can open the configured workbook, and the server is running in an interactive Windows desktop session.
- If Selenium validation fails, confirm Chrome is installed and available to Selenium Manager.
