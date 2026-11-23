import os
import sys
import json
import socket
import getpass
import platform
import threading
from datetime import datetime
from pathlib import Path
import urllib.request
import urllib.error

# URL Webhook mặc định nếu chưa cấu hình nơi khác
DEFAULT_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbwl0acXpTU-IPBgJwRwMHHE9ieinyhSjwnA-og0b3XvOyaBHjprIFPTlDEA4Pi1bJfanw/exec"

def _find_config_file() -> Path | None:
    candidates = [
        Path(__file__).resolve().parent.parent.parent / "google_sheet_config.json",
        Path.cwd() / "google_sheet_config.json",
        Path(sys.executable).parent / "google_sheet_config.json",
        Path(getattr(sys, "_MEIPASS", ".")) / "google_sheet_config.json",
    ]
    for p in candidates:
        if p.is_file():
            return p
    return candidates[0]

CONFIG_FILE = _find_config_file() or (Path(__file__).resolve().parent.parent.parent / "google_sheet_config.json")

def get_google_sheet_url() -> str:
    """Lấy URL Google Sheet Webhook từ Streamlit secrets, biến môi trường, file config hoặc mặc định."""
    # 1. Từ Streamlit secrets nếu đang chạy trên Streamlit Cloud
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "GOOGLE_SHEET_WEBHOOK_URL" in st.secrets:
            sec_url = str(st.secrets["GOOGLE_SHEET_WEBHOOK_URL"]).strip()
            if sec_url:
                return sec_url
    except Exception:
        pass

    # 2. Từ biến môi trường
    env_url = os.environ.get("GOOGLE_SHEET_WEBHOOK_URL", "").strip()
    if env_url:
        return env_url
    
    # 3. Từ file cấu hình local google_sheet_config.json
    cfg = _find_config_file()
    if cfg and cfg.is_file():
        try:
            with open(cfg, "r", encoding="utf-8") as f:
                data = json.load(f)
                url = data.get("webhook_url", "").strip()
                if url:
                    return url
        except Exception:
            pass
            
    # 4. Mặc định
    return DEFAULT_WEBHOOK_URL

def set_google_sheet_url(url: str) -> None:
    """Lưu URL Google Sheet Webhook vào file config."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"webhook_url": url.strip()}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _send_log_payload(payload: dict) -> None:
    """Gửi dữ liệu nhật ký tới Google Sheet Webhook (chạy ngầm)."""
    url = get_google_sheet_url()
    if not url:
        return
    try:
        data_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=data_bytes,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=5) as response:
            pass
    except Exception:
        # Chạy ngầm, không làm gián đoạn hay báo lỗi ứng dụng
        pass

def get_local_device_info() -> str:
    """Lấy tên máy tính và model phần cứng (vd: SG-KI-VINHPH-UI (21JK006HVA))."""
    name = os.environ.get("COMPUTERNAME") or socket.gethostname() or "Máy tính"
    model = ""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DESCRIPTION\System\BIOS") as key:
            model = winreg.QueryValueEx(key, "SystemProductName")[0].strip()
    except Exception:
        pass
    if model and model.lower() not in ("system product name", "to be filled by o.e.m.", "system manufacturer"):
        return f"{name} ({model})"
    return name

def log_system_access(
    app_type: str = "Website",
    action: str = "Truy cập hệ thống",
    user_info: str = "",
    device_info: str = "",
    details: str = "",
) -> None:
    """
    Ghi lại lịch sử đăng nhập/sử dụng vào Google Sheet bất đồng bộ.
    """
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if not device_info:
        try:
            device_info = get_local_device_info()
        except Exception:
            device_info = "Máy khách"

    if not user_info:
        try:
            user_info = getpass.getuser()
        except Exception:
            user_info = "Ẩn danh"

    payload = {
        "timestamp": now_str,
        "app_type": app_type,
        "device": device_info,
        "user": user_info,
        "action": action,
        "details": details,
    }

    # Bắn request trong thread riêng biệt để không chặn UI
    thread = threading.Thread(target=_send_log_payload, args=(payload,), daemon=True)
    thread.start()
