import threading
import time
import traceback
import re

from fastapi import FastAPI, HTTPException


app = FastAPI(title="GST Extraction API", version="1.0.0")
GST_REGEX = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")
state_lock = threading.Lock()
job_state = {
    "running": False,
    "last_started_at": None,
    "last_finished_at": None,
    "last_status": "idle",
    "last_error": None,
    "last_gst_number": None,
    "last_result": None,
}


def _set_state(**kwargs):
    with state_lock:
        job_state.update(kwargs)


def _run_job(gst_number):
    _set_state(
        running=True,
        last_started_at=time.strftime("%Y-%m-%d %H:%M:%S"),
        last_finished_at=None,
        last_status="running",
        last_error=None,
        last_gst_number=gst_number,
        last_result=None,
    )
    try:
        import main as gst_main

        results = gst_main.main([gst_number], save_to_excel=False)
        _set_state(last_status="completed", last_result=results)
    except Exception:
        _set_state(last_status="failed", last_error=traceback.format_exc())
    finally:
        _set_state(running=False, last_finished_at=time.strftime("%Y-%m-%d %H:%M:%S"))


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/status")
def status():
    with state_lock:
        return dict(job_state)


@app.post("/extract")
def extract(gst_number: str):
    gst_number = (gst_number or "").strip().upper().replace(" ", "")
    if not gst_number:
        raise HTTPException(status_code=400, detail="gst_number is required.")
    if not GST_REGEX.fullmatch(gst_number):
        return {
            "GST Number": gst_number,
            "Trade Name": "",
            "Taxpayer Type": "",
            "GSTIN / UIN  Status": "Invalid",
        }

    with state_lock:
        if job_state["running"]:
            raise HTTPException(status_code=409, detail="Extraction already running.")

    _run_job(gst_number)
    with state_lock:
        if job_state.get("last_status") == "completed":
            results = job_state.get("last_result") or []
            if results:
                return results[0]
            return {
                "GST Number": gst_number,
                "Trade Name": "",
                "Taxpayer Type": "",
                "GSTIN / UIN  Status": "Error",
            }
        if job_state.get("last_status") == "failed":
            raise HTTPException(status_code=500, detail="Extraction failed. Check /status for details.")
        raise HTTPException(status_code=500, detail="Unexpected extraction state.")
