import os
import re
import subprocess
import time
from pathlib import Path
from typing import Any


def _normalize_path(value: Path | str) -> Path:
    return Path(str(value or "")).expanduser()


def _resolve_path_text(value: Path | str) -> str:
    try:
        return str(_normalize_path(value).resolve())
    except Exception:
        return str(_normalize_path(value))


def _to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"1", "true", "yes", "y", "on"}


def _start_process_maximized(executable_path: Path) -> None:
    startupinfo = None
    if os.name == "nt":
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 3
        except Exception:
            startupinfo = None
    subprocess.Popen(
        [str(executable_path)],
        cwd=str(executable_path.parent),
        startupinfo=startupinfo,
    )


def _empty_post_sync_result(
    *,
    vendor_upload_path: Path | str,
    sap_logon_path: Path | str | None = None,
    sap_server_name: str = "",
    sap_username: str = "",
    sap_client: str = "",
    sap_transaction_code: str = "",
) -> dict[str, Any]:
    normalized_vendor_upload_path = _normalize_path(vendor_upload_path)
    normalized_sap_logon_path = _normalize_path(sap_logon_path or "")
    return {
        "vendorUploadOpened": False,
        "vendorUploadOpenPath": str(normalized_vendor_upload_path),
        "vendorUploadOpenError": None,
        "vendorUploadMinimized": False,
        "vendorUploadMinimizePath": str(normalized_vendor_upload_path),
        "vendorUploadMinimizeError": None,
        "vendorUploadClosed": False,
        "vendorUploadClosePath": str(normalized_vendor_upload_path),
        "vendorUploadCloseError": None,
        "sapLogonStarted": False,
        "sapLogonPath": str(normalized_sap_logon_path),
        "sapLogonError": None,
        "sapServerName": str(sap_server_name or ""),
        "sapUser": str(sap_username or ""),
        "sapClient": str(sap_client or ""),
        "sapAutoLogonSuccess": False,
        "sapAutoLogonError": None,
        "sapSecurityPopupsConfirmed": 0,
        "sapTransactionCode": str(sap_transaction_code or ""),
        "sapTransactionWaitSeconds": 0,
        "sapExcelMinimizeDelaySeconds": 0,
        "sapTransactionSubmitted": False,
        "sapTransactionError": None,
        "sapVendorCode": None,
        "sapVendorCodePopupFound": False,
        "sapVendorCodePopupTitle": None,
        "sapVendorCodePopupText": None,
        "sapVendorCodePopupConfirmed": False,
    }


def _coerce_timeout_seconds(value: Any, default: int = 120) -> int:
    try:
        return max(5, int(value))
    except (TypeError, ValueError):
        return default


def _coerce_wait_seconds(value: Any, default: int = 20) -> int:
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return default


def _import_sap_com_modules() -> tuple[Any, Any]:
    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client as win32com  # type: ignore[import-untyped]
    except Exception as exc:
        raise RuntimeError(f"pywin32 unavailable for SAP GUI automation: {exc}") from exc
    return pythoncom, win32com


def _get_collection_count(collection: Any) -> int:
    try:
        return int(getattr(collection, "Count", 0) or 0)
    except Exception:
        return 0


def _get_child(collection: Any, index: int) -> Any:
    try:
        return collection(index)
    except Exception:
        return collection.Item(index)


def _wait_until_ready(session: Any, timeout_seconds: int) -> None:
    deadline = time.time() + _coerce_timeout_seconds(timeout_seconds)
    while True:
        try:
            if not bool(getattr(session, "Busy", False)):
                return
        except Exception:
            return
        if time.time() >= deadline:
            raise RuntimeError("Timed out waiting for SAP session to become ready")
        time.sleep(0.25)


def _wait_for_sap_application(win32com: Any, timeout_seconds: int) -> Any:
    deadline = time.time() + _coerce_timeout_seconds(timeout_seconds)
    last_error: Exception | None = None

    while time.time() < deadline:
        try:
            sap_gui_auto = win32com.GetObject("SAPGUI")
            return sap_gui_auto.GetScriptingEngine
        except Exception as exc:
            last_error = exc
            time.sleep(1)

    if last_error:
        raise RuntimeError(f"Timed out waiting for SAP GUI scripting engine: {last_error}") from last_error
    raise RuntimeError("Timed out waiting for SAP GUI scripting engine")


def _start_sap_logon_if_needed(
    *,
    win32com: Any,
    sap_logon_path: Path | str,
    auto_launch_logon: bool,
    timeout_seconds: int,
) -> dict[str, Any]:
    target_path = _normalize_path(sap_logon_path)
    result = {
        "sapLogonStarted": False,
        "sapLogonPath": _resolve_path_text(target_path),
        "sapLogonError": None,
    }

    try:
        _wait_for_sap_application(win32com, 2)
        return result
    except Exception:
        pass

    if not _to_bool(auto_launch_logon):
        result["sapLogonError"] = "SAP GUI scripting engine is not available and auto launch is disabled"
        return result

    if not target_path.exists() or not target_path.is_file():
        result["sapLogonError"] = f"SAP Logon executable not found: {target_path}"
        return result

    try:
        _start_process_maximized(target_path)
        result["sapLogonStarted"] = True
        _wait_for_sap_application(win32com, timeout_seconds)
    except Exception as exc:
        result["sapLogonError"] = str(exc)

    return result


def _open_sap_connection(application: Any, sap_server_name: str, timeout_seconds: int) -> Any:
    server_name = str(sap_server_name or "").strip()
    if server_name:
        try:
            connection = application.OpenConnection(server_name, True)
            if connection is not None:
                return connection
        except Exception as exc:
            raise RuntimeError(f"Unable to open SAP connection '{server_name}': {exc}") from exc

    deadline = time.time() + _coerce_timeout_seconds(timeout_seconds)
    while time.time() < deadline:
        connection_count = _get_collection_count(application.Children)
        if connection_count > 0:
            return _get_child(application.Children, connection_count - 1)
        time.sleep(1)

    raise RuntimeError("No SAP GUI connection is available")


def _wait_for_sap_session(connection: Any, timeout_seconds: int) -> Any:
    deadline = time.time() + _coerce_timeout_seconds(timeout_seconds)
    while time.time() < deadline:
        session_count = _get_collection_count(connection.Children)
        if session_count > 0:
            session = _get_child(connection.Children, 0)
            _wait_until_ready(session, timeout_seconds)
            return session
        time.sleep(1)
    raise RuntimeError("Timed out waiting for SAP session")


def _try_set_text(session: Any, element_id: str, value: Any) -> bool:
    text = str(value or "")
    if not text:
        return False
    try:
        field = session.findById(element_id)
        field.text = text
        return True
    except Exception:
        return False


def _try_press(session: Any, element_id: str) -> bool:
    try:
        session.findById(element_id).press()
        return True
    except Exception:
        return False


def _find_optional(session: Any, element_id: str) -> Any | None:
    try:
        return session.findById(element_id)
    except Exception:
        return None


def _window_exists(session: Any, window_id: str) -> bool:
    return _find_optional(session, window_id) is not None


def _show_and_maximize_sap_window(session: Any) -> bool:
    window = _find_optional(session, "wnd[0]")
    sap_title = ""
    maximized = False

    if window is not None:
        try:
            sap_title = str(getattr(window, "Text", "") or "").strip()
        except Exception:
            sap_title = ""

        for method_name in ("maximize", "Maximize"):
            try:
                method = getattr(window, method_name)
                if callable(method):
                    method()
                    maximized = True
                    break
            except Exception:
                continue

        for focus_method_name in ("setFocus", "SetFocus"):
            try:
                focus_method = getattr(window, focus_method_name)
                if callable(focus_method):
                    focus_method()
                    break
            except Exception:
                continue

    try:
        import win32con  # type: ignore[import-untyped]
        import win32gui  # type: ignore[import-untyped]
    except Exception:
        return maximized

    candidate_hwnds: list[int] = []

    def collect_sap_windows(hwnd: int, _: Any) -> None:
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title_text = str(win32gui.GetWindowText(hwnd) or "").strip()
            class_name = str(win32gui.GetClassName(hwnd) or "").strip().lower()
        except Exception:
            return

        if sap_title and title_text and sap_title.lower() in title_text.lower():
            candidate_hwnds.append(hwnd)
            return

        if "sap" in class_name or "sap" in title_text.lower():
            candidate_hwnds.append(hwnd)

    try:
        win32gui.EnumWindows(collect_sap_windows, None)
    except Exception:
        return maximized

    for hwnd in candidate_hwnds:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            maximized = True
            break
        except Exception:
            continue

    return maximized


def _confirm_sap_popup(session: Any) -> bool:
    popup_button_ids = (
        "wnd[1]/tbar[0]/btn[0]",
        "wnd[1]/usr/btnSPOP-OPTION1",
        "wnd[1]/usr/btnBUTTON_1",
        "wnd[1]/usr/btnBUTTON_2",
    )
    for button_id in popup_button_ids:
        if _try_press(session, button_id):
            return True

    try:
        session.findById("wnd[1]").sendVKey(0)
        return True
    except Exception:
        pass

    try:
        session.findById("wnd[1]").close()
        return True
    except Exception:
        return False


def _confirm_initial_sap_popups(session: Any, timeout_seconds: int) -> int:
    deadline = time.time() + min(_coerce_timeout_seconds(timeout_seconds), 20)
    confirmed = 0
    while time.time() < deadline and _window_exists(session, "wnd[1]"):
        if not _confirm_sap_popup(session):
            break
        confirmed += 1
        time.sleep(1)
    return confirmed


def _log_in_to_sap_session(
    session: Any,
    *,
    sap_username: str,
    sap_password: str,
    sap_client: str,
    sap_language: str,
    timeout_seconds: int,
) -> int:
    _wait_until_ready(session, timeout_seconds)

    if sap_client:
        _try_set_text(session, "wnd[0]/usr/txtRSYST-MANDT", sap_client)
    if sap_language:
        _try_set_text(session, "wnd[0]/usr/txtRSYST-LANGU", sap_language)

    username_set = _try_set_text(session, "wnd[0]/usr/txtRSYST-BNAME", sap_username)
    password_set = _try_set_text(session, "wnd[0]/usr/pwdRSYST-BCODE", sap_password)
    if not username_set or not password_set:
        if _find_optional(session, "wnd[0]/tbar[0]/okcd") is not None:
            return _confirm_initial_sap_popups(session, timeout_seconds)
        raise RuntimeError("SAP login fields were not found; the session may already be logged in or on an unexpected screen")

    password_field = session.findById("wnd[0]/usr/pwdRSYST-BCODE")
    try:
        password_field.setFocus()
        password_field.caretPosition = len(str(sap_password or ""))
    except Exception:
        pass

    session.findById("wnd[0]").sendVKey(0)
    _wait_until_ready(session, timeout_seconds)
    return _confirm_initial_sap_popups(session, timeout_seconds)


def _extract_text_from_sap_component(component: Any, collected: list[str], depth: int = 0) -> None:
    if depth > 5:
        return

    for attribute_name in ("Text", "text"):
        try:
            value = getattr(component, attribute_name)
            text = str(value or "").strip()
            if text:
                collected.append(text)
                break
        except Exception:
            pass

    try:
        children = component.Children
        child_count = _get_collection_count(children)
    except Exception:
        child_count = 0

    for index in range(child_count):
        try:
            _extract_text_from_sap_component(_get_child(children, index), collected, depth + 1)
        except Exception:
            continue


def _get_sap_popup_details(session: Any) -> dict[str, Any]:
    popup = _find_optional(session, "wnd[1]")
    if popup is None:
        return {
            "sapVendorCodePopupFound": False,
            "sapVendorCodePopupTitle": None,
            "sapVendorCodePopupText": None,
        }

    title = ""
    try:
        title = str(getattr(popup, "Text", "") or "").strip()
    except Exception:
        title = ""

    collected: list[str] = []
    for element_id in (
        "wnd[1]/usr/txtMESSTXT1",
        "wnd[1]/usr/txtMESSTXT2",
        "wnd[1]/usr/txtMESSTXT3",
        "wnd[1]/usr/txtMESSTXT4",
    ):
        element = _find_optional(session, element_id)
        if element is not None:
            _extract_text_from_sap_component(element, collected)

    if not collected:
        _extract_text_from_sap_component(popup, collected)

    unique_parts: list[str] = []
    seen: set[str] = set()
    for item in collected:
        normalized = " ".join(str(item or "").split())
        if normalized and normalized not in seen:
            seen.add(normalized)
            unique_parts.append(normalized)

    return {
        "sapVendorCodePopupFound": True,
        "sapVendorCodePopupTitle": title or None,
        "sapVendorCodePopupText": " | ".join(unique_parts) or None,
    }


def _wait_for_sap_popup_details(session: Any, timeout_seconds: int) -> dict[str, Any]:
    deadline = time.time() + _coerce_timeout_seconds(timeout_seconds, default=30)
    while time.time() < deadline:
        popup_details = _get_sap_popup_details(session)
        if popup_details.get("sapVendorCodePopupFound"):
            return popup_details
        time.sleep(1)
    return _get_sap_popup_details(session)


def _extract_vendor_code_from_popup_text(text: str | None) -> str | None:
    value = str(text or "").strip()
    if not value:
        return None

    patterns = (
        r"(?:vendor|code|number|created)\D{0,40}([A-Z0-9]{4,20})",
        r"\b([0-9]{5,12})\b",
    )
    for pattern in patterns:
        match = re.search(pattern, value, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def _capture_excel_state(workbook_path: Path | str, *, manage_com: bool = True) -> dict[str, Any]:
    target_path = _normalize_path(workbook_path)
    target_full_path = _resolve_path_text(target_path).lower()
    result: dict[str, Any] = {
        "path": _resolve_path_text(target_path),
        "appOpen": False,
        "workbookCount": 0,
        "windowCount": 0,
        "targetWorkbookOpen": False,
        "error": None,
    }

    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client as win32com  # type: ignore[import-untyped]
    except Exception as exc:
        result["error"] = f"pywin32 unavailable for Excel state check: {exc}"
        return result

    if manage_com:
        pythoncom.CoInitialize()
    try:
        try:
            excel_app = win32com.GetActiveObject("Excel.Application")
        except Exception:
            return result

        result["appOpen"] = True
        try:
            result["workbookCount"] = int(getattr(excel_app.Workbooks, "Count", 0) or 0)
        except Exception:
            result["workbookCount"] = 0

        try:
            result["windowCount"] = int(getattr(excel_app.Windows, "Count", 0) or 0)
        except Exception:
            result["windowCount"] = 0

        workbook_count = int(result["workbookCount"] or 0)
        for index in range(workbook_count, 0, -1):
            try:
                workbook = excel_app.Workbooks.Item(index)
                workbook_path_text = _resolve_path_text(getattr(workbook, "FullName", "")).lower()
                if workbook_path_text == target_full_path:
                    result["targetWorkbookOpen"] = True
                    break
            except Exception:
                continue
    except Exception as exc:
        result["error"] = str(exc)
    finally:
        if manage_com:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return result


def _wait_for_excel_to_open(
    workbook_path: Path | str,
    timeout_seconds: int,
    *,
    manage_com: bool = True,
) -> dict[str, Any]:
    deadline = time.time() + _coerce_timeout_seconds(timeout_seconds, default=30)
    last_state = _capture_excel_state(workbook_path, manage_com=manage_com)

    while time.time() < deadline:
        current_state = _capture_excel_state(workbook_path, manage_com=manage_com)
        last_state = current_state
        if (
            bool(current_state.get("targetWorkbookOpen"))
            or int(current_state.get("windowCount") or 0) > 0
            or int(current_state.get("workbookCount") or 0) > 0
        ):
            current_state["excelAvailable"] = True
            return current_state
        time.sleep(0.5)

    last_state["excelAvailable"] = False
    return last_state


def _run_sap_transaction(
    session: Any,
    *,
    vendor_upload_path: Path | str,
    close_vendor_upload_after_sync: bool,
    minimize_vendor_upload_after_transaction: bool,
    sap_transaction_code: str,
    timeout_seconds: int,
    sap_transaction_wait_seconds: int = 20,
    sap_excel_minimize_delay_seconds: int = 3,
) -> dict[str, Any]:
    transaction_code = str(sap_transaction_code or "").strip()
    wait_seconds = _coerce_wait_seconds(sap_transaction_wait_seconds)
    minimize_delay_seconds = _coerce_wait_seconds(sap_excel_minimize_delay_seconds, default=3)
    if not transaction_code:
        return {
            "vendorUploadMinimized": False,
            "vendorUploadMinimizePath": str(_normalize_path(vendor_upload_path)),
            "vendorUploadMinimizeError": None,
            "vendorUploadClosed": False,
            "vendorUploadClosePath": str(_normalize_path(vendor_upload_path)),
            "vendorUploadCloseError": None,
            "sapTransactionWaitSeconds": wait_seconds,
            "sapExcelMinimizeDelaySeconds": minimize_delay_seconds,
            "sapTransactionSubmitted": False,
            "sapTransactionError": "SAP transaction code is empty",
            "sapVendorCode": None,
            "sapVendorCodePopupFound": False,
            "sapVendorCodePopupTitle": None,
            "sapVendorCodePopupText": None,
            "sapVendorCodePopupConfirmed": False,
        }

    try:
        _wait_until_ready(session, timeout_seconds)
        ok_code = session.findById("wnd[0]/tbar[0]/okcd")
        ok_code.text = transaction_code
        session.findById("wnd[0]").sendVKey(0)
        _wait_until_ready(session, timeout_seconds)

        if wait_seconds > 0:
            time.sleep(wait_seconds)

        workbook_minimize_result: dict[str, Any] | None = None
        if _to_bool(minimize_vendor_upload_after_transaction):
            excel_state = _wait_for_excel_to_open(
                vendor_upload_path,
                timeout_seconds,
                manage_com=False,
            )
            if bool(excel_state.get("excelAvailable")):
                if minimize_delay_seconds > 0:
                    time.sleep(minimize_delay_seconds)
                workbook_minimize_result = minimize_vendor_upload_workbook(
                    vendor_upload_path,
                    manage_com=False,
                )
            else:
                workbook_minimize_result = {
                    "path": str(excel_state.get("path") or _resolve_path_text(vendor_upload_path)),
                    "minimized": False,
                    "alreadyMinimized": False,
                    "alreadyClosed": True,
                    "error": excel_state.get("error") or "Excel did not open after SAP transaction",
                }

        popup_details = _wait_for_sap_popup_details(session, timeout_seconds)

        result = {
            "vendorUploadMinimized": False,
            "vendorUploadMinimizePath": str(_normalize_path(vendor_upload_path)),
            "vendorUploadMinimizeError": None,
            "sapTransactionSubmitted": True,
            "sapTransactionError": None,
            "sapTransactionWaitSeconds": wait_seconds,
            "sapExcelMinimizeDelaySeconds": minimize_delay_seconds,
            "sapVendorCode": _extract_vendor_code_from_popup_text(
                str(popup_details.get("sapVendorCodePopupText") or "")
            ),
            "sapVendorCodePopupConfirmed": False,
            **popup_details,
        }
        if _to_bool(close_vendor_upload_after_sync):
            result.update(
                {
                    "vendorUploadClosed": False,
                    "vendorUploadClosePath": str(_normalize_path(vendor_upload_path)),
                    "vendorUploadCloseError": None,
                }
            )
        if workbook_minimize_result is not None:
            result.update(
                {
                    "vendorUploadMinimized": bool(
                        workbook_minimize_result.get("minimized")
                        or workbook_minimize_result.get("alreadyMinimized")
                    ),
                    "vendorUploadMinimizePath": str(
                        workbook_minimize_result.get("path") or str(_normalize_path(vendor_upload_path))
                    ),
                    "vendorUploadMinimizeError": workbook_minimize_result.get("error"),
                }
            )
        return result
    except Exception as exc:
        return {
            "vendorUploadMinimized": False,
            "vendorUploadMinimizePath": str(_normalize_path(vendor_upload_path)),
            "vendorUploadMinimizeError": None,
            "vendorUploadClosed": False,
            "vendorUploadClosePath": str(_normalize_path(vendor_upload_path)),
            "vendorUploadCloseError": None,
            "sapTransactionWaitSeconds": wait_seconds,
            "sapExcelMinimizeDelaySeconds": minimize_delay_seconds,
            "sapTransactionSubmitted": False,
            "sapTransactionError": str(exc),
            "sapVendorCode": None,
            "sapVendorCodePopupFound": False,
            "sapVendorCodePopupTitle": None,
            "sapVendorCodePopupText": None,
            "sapVendorCodePopupConfirmed": False,
        }


def connect_to_sap() -> Any:
    """Connect to the active SAP GUI session and return the first session."""
    _, win32com = _import_sap_com_modules()
    application = _wait_for_sap_application(win32com, 30)
    connection = _open_sap_connection(application, "", 30)
    return _wait_for_sap_session(connection, 30)


def run_zvend_up() -> dict[str, Any]:
    return run_post_vendor_upload_process(
        vendor_upload_path="",
        open_vendor_upload_after_sync=False,
        close_vendor_upload_after_sync=False,
        auto_launch_logon=False,
        auto_logon_after_launch=True,
        sap_server_name=os.getenv("SAP_SERVER_NAME", ""),
        sap_username=os.getenv("SAP_LOGON_USERNAME", ""),
        sap_password=os.getenv("SAP_LOGON_PASSWORD", ""),
        sap_client=os.getenv("SAP_LOGON_CLIENT", ""),
        sap_language=os.getenv("SAP_LOGON_LANGUAGE", ""),
        sap_wait_timeout_seconds=120,
        sap_transaction_wait_seconds=0,
        sap_excel_minimize_delay_seconds=3,
        sap_transaction_code=os.getenv("SAP_VENDOR_UPLOAD_TRANSACTION_CODE", "ZVEND_UP") or "ZVEND_UP",
    )


def open_vendor_upload_workbook(workbook_path: Path | str) -> dict[str, Any]:
    target_path = _normalize_path(workbook_path)
    result: dict[str, Any] = {
        "path": _resolve_path_text(target_path),
        "started": False,
        "error": None,
    }

    if not target_path.exists() or not target_path.is_file():
        result["error"] = f"Vendor upload workbook not found: {target_path}"
        return result

    try:
        if hasattr(os, "startfile"):
            os.startfile(str(target_path))  # type: ignore[attr-defined]
        else:
            subprocess.Popen([str(target_path)], cwd=str(target_path.parent))
        result["started"] = True
    except Exception as exc:
        result["error"] = str(exc)

    return result


def close_vendor_upload_workbook(workbook_path: Path | str, *, manage_com: bool = True) -> dict[str, Any]:
    target_path = _normalize_path(workbook_path)
    target_full_path = _resolve_path_text(target_path).lower()
    result: dict[str, Any] = {
        "path": _resolve_path_text(target_path),
        "closed": False,
        "alreadyClosed": False,
        "excelQuit": False,
        "error": None,
    }

    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client as win32com  # type: ignore[import-untyped]
    except Exception as exc:
        result["error"] = f"pywin32 unavailable for Excel close: {exc}"
        return result

    if manage_com:
        pythoncom.CoInitialize()
    try:
        try:
            excel_app = win32com.GetActiveObject("Excel.Application")
        except Exception:
            result["alreadyClosed"] = True
            return result

        workbook_count = int(getattr(excel_app.Workbooks, "Count", 0) or 0)
        for index in range(workbook_count, 0, -1):
            try:
                workbook = excel_app.Workbooks.Item(index)
                workbook_path_text = _resolve_path_text(getattr(workbook, "FullName", "")).lower()
                if workbook_path_text == target_full_path:
                    workbook.Close(SaveChanges=False)
                    result["closed"] = True
                    try:
                        if int(getattr(excel_app.Workbooks, "Count", 0) or 0) == 0:
                            excel_app.Quit()
                            result["excelQuit"] = True
                    except Exception:
                        pass
                    break
            except Exception:
                continue

        if not result["closed"]:
            result["alreadyClosed"] = True
    except Exception as exc:
        result["error"] = str(exc)
    finally:
        if manage_com:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return result


def minimize_vendor_upload_workbook(workbook_path: Path | str, *, manage_com: bool = True) -> dict[str, Any]:
    target_path = _normalize_path(workbook_path)
    result: dict[str, Any] = {
        "path": _resolve_path_text(target_path),
        "minimized": False,
        "alreadyMinimized": False,
        "alreadyClosed": False,
        "error": None,
    }

    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client as win32com  # type: ignore[import-untyped]
        import win32con  # type: ignore[import-untyped]
        import win32gui  # type: ignore[import-untyped]
    except Exception as exc:
        result["error"] = f"pywin32 unavailable for Excel minimize: {exc}"
        return result

    if manage_com:
        pythoncom.CoInitialize()
    try:
        try:
            excel_app = win32com.GetActiveObject("Excel.Application")
        except Exception:
            result["alreadyClosed"] = True
            return result

        xl_minimized = -4140
        minimized_any = False

        try:
            excel_app.Visible = True
        except Exception:
            pass

        try:
            current_state = int(getattr(excel_app, "WindowState", 0) or 0)
        except Exception:
            current_state = 0

        if current_state != xl_minimized:
            try:
                excel_app.WindowState = xl_minimized
                minimized_any = True
            except Exception:
                pass

        def minimize_excel_window(hwnd: int, _: Any) -> None:
            nonlocal minimized_any
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32gui.GetClassName(hwnd) != "XLMAIN":
                    return
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                minimized_any = True
            except Exception:
                return

        try:
            win32gui.EnumWindows(minimize_excel_window, None)
        except Exception:
            pass

        if minimized_any:
            result["minimized"] = True
        else:
            result["alreadyMinimized"] = True
    except Exception as exc:
        result["error"] = str(exc)
    finally:
        if manage_com:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return result


def _kill_process_images(process_names: list[str]) -> dict[str, Any]:
    result: dict[str, Any] = {
        "killed": False,
        "alreadyClosed": False,
        "error": None,
        "attempted": list(process_names),
        "outputs": [],
    }

    details: list[str] = []
    killed_any = False
    meaningful_errors: list[str] = []
    for process_name in process_names:
        image_name = str(process_name or "").strip()
        if not image_name:
            continue
        try:
            completed = subprocess.run(
                ["taskkill", "/F", "/T", "/IM", image_name],
                capture_output=True,
                text=True,
                timeout=15,
            )
        except Exception as exc:
            meaningful_errors.append(f"{image_name}: {exc}")
            continue

        stdout_text = str(completed.stdout or "").strip()
        stderr_text = str(completed.stderr or "").strip()
        combined = " | ".join(part for part in (stdout_text, stderr_text) if part)
        if combined:
            details.append(f"{image_name}: {combined}")

        if completed.returncode == 0:
            killed_any = True
            continue

        lowered = combined.lower()
        if "not found" in lowered or "no running instance" in lowered or "could not be found" in lowered:
            continue
        if combined:
            meaningful_errors.append(f"{image_name}: {combined}")
        else:
            meaningful_errors.append(f"{image_name}: taskkill exited with code {completed.returncode}")

    result["killed"] = killed_any
    result["alreadyClosed"] = not killed_any and not meaningful_errors
    result["outputs"] = details
    if meaningful_errors:
        result["error"] = "; ".join(meaningful_errors)
    return result


def kill_excel_processes() -> dict[str, Any]:
    return _kill_process_images(["EXCEL.EXE"])


def kill_sap_processes() -> dict[str, Any]:
    return _kill_process_images([
        "saplogon.exe",
        "saplgpad.exe",
        "sapfewse.exe",
        "sapshcut.exe",
    ])


def close_active_sap_session(*, timeout_seconds: int = 15, manage_com: bool = True) -> dict[str, Any]:
    _ = timeout_seconds
    _ = manage_com
    kill_result = kill_sap_processes()
    return {
        "closed": bool(kill_result.get("killed")),
        "alreadyClosed": bool(kill_result.get("alreadyClosed")),
        "error": kill_result.get("error"),
        "outputs": kill_result.get("outputs"),
        "attempted": kill_result.get("attempted"),
    }


def run_post_vendor_upload_process(
    *,
    vendor_upload_path: Path | str,
    open_vendor_upload_after_sync: bool,
    close_vendor_upload_after_sync: bool,
    sap_logon_path: Path | str | None = None,
    auto_launch_logon: bool = False,
    auto_logon_after_launch: bool = False,
    sap_server_name: str = "",
    sap_username: str = "",
    sap_password: str = "",
    sap_client: str = "",
    sap_language: str = "",
    sap_wait_timeout_seconds: int = 120,
    sap_transaction_wait_seconds: int = 20,
    sap_excel_minimize_delay_seconds: int = 3,
    sap_transaction_code: str = "",
    **_: Any,
) -> dict[str, Any]:
    timeout_seconds = _coerce_timeout_seconds(sap_wait_timeout_seconds)
    transaction_wait_seconds = _coerce_wait_seconds(sap_transaction_wait_seconds)
    minimize_delay_seconds = _coerce_wait_seconds(sap_excel_minimize_delay_seconds, default=3)
    output = _empty_post_sync_result(
        vendor_upload_path=vendor_upload_path,
        sap_logon_path=sap_logon_path,
        sap_server_name=sap_server_name,
        sap_username=sap_username,
        sap_client=sap_client,
        sap_transaction_code=sap_transaction_code,
    )
    output["sapTransactionWaitSeconds"] = transaction_wait_seconds
    output["sapExcelMinimizeDelaySeconds"] = minimize_delay_seconds
    sap_automation_enabled = _to_bool(auto_launch_logon) or _to_bool(auto_logon_after_launch)

    if _to_bool(close_vendor_upload_after_sync) and not sap_automation_enabled:
        workbook_close_result = close_vendor_upload_workbook(vendor_upload_path)
        output["vendorUploadClosed"] = bool(
            workbook_close_result.get("closed") or workbook_close_result.get("alreadyClosed")
        )
        output["vendorUploadClosePath"] = str(workbook_close_result.get("path") or output["vendorUploadClosePath"])
        output["vendorUploadCloseError"] = workbook_close_result.get("error")

    if _to_bool(open_vendor_upload_after_sync):
        workbook_result = open_vendor_upload_workbook(vendor_upload_path)
        output["vendorUploadOpened"] = bool(workbook_result.get("started"))
        output["vendorUploadOpenPath"] = str(workbook_result.get("path") or output["vendorUploadOpenPath"])
        output["vendorUploadOpenError"] = workbook_result.get("error")

    if not sap_automation_enabled:
        return output

    pythoncom = None
    try:
        pythoncom, win32com = _import_sap_com_modules()
        pythoncom.CoInitialize()

        launch_result = _start_sap_logon_if_needed(
            win32com=win32com,
            sap_logon_path=sap_logon_path or "",
            auto_launch_logon=auto_launch_logon,
            timeout_seconds=timeout_seconds,
        )
        output.update(launch_result)
        if output.get("sapLogonError"):
            return output

        application = _wait_for_sap_application(win32com, timeout_seconds)
        connection = _open_sap_connection(application, sap_server_name, timeout_seconds)
        session = _wait_for_sap_session(connection, timeout_seconds)
        _show_and_maximize_sap_window(session)

        if _to_bool(auto_logon_after_launch):
            try:
                confirmed = _log_in_to_sap_session(
                    session,
                    sap_username=sap_username,
                    sap_password=sap_password,
                    sap_client=sap_client,
                    sap_language=sap_language,
                    timeout_seconds=timeout_seconds,
                )
                output["sapAutoLogonSuccess"] = True
                output["sapSecurityPopupsConfirmed"] = confirmed
                _show_and_maximize_sap_window(session)
            except Exception as exc:
                output["sapAutoLogonError"] = str(exc)
                return output

            transaction_result = _run_sap_transaction(
                session,
                vendor_upload_path=vendor_upload_path,
                close_vendor_upload_after_sync=close_vendor_upload_after_sync,
                minimize_vendor_upload_after_transaction=True,
                sap_transaction_code=sap_transaction_code,
                timeout_seconds=timeout_seconds,
                sap_transaction_wait_seconds=transaction_wait_seconds,
                sap_excel_minimize_delay_seconds=minimize_delay_seconds,
            )
            output.update(transaction_result)
    except Exception as exc:
        output["sapAutoLogonError"] = str(exc)
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return output


if __name__ == "__main__":
    result = run_zvend_up()
    print(result)
