# FastAPI Setup and Run Instructions

## 1) Open project folder

```powershell
cd "c:\Users\Maulik Khunt\Desktop\API Codes\GST Extraction Process"
```

## 2) Create and activate virtual environment

```powershell
python -m venv venv
venv\Scripts\activate
```

## 3) Install required packages

```powershell
pip install fastapi uvicorn openpyxl selenium numpy Pillow opencv-python easyocr transformers torch
```

## 4) Run API server

```powershell
uvicorn api:app --host 0.0.0.0 --port 8000 --reload
```

## 5) Open API docs

- Swagger UI: `http://127.0.0.1:8000/docs`

## 6) API endpoints

- `GET /health`
- `GET /status`
- `POST /extract?gst_number=<GSTIN>`

Example:

```http
POST http://127.0.0.1:8000/extract?gst_number=24AAABC4799H1Z1
```

## 7) Notes

- API processes one GST number per request.
- `/extract` runs synchronously and returns result in response.
- Use `/status` to view last run details.
- API does not use Excel input/output; extraction runs directly from `gst_number`.
