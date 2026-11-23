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

# Đường dẫn file cấu hình URL Google Sheet Webhook
CONFIG_FILE = Path(__file__).resolve().parent.parent.parent / "google_sheet_config.json"

def get_google_sheet_url() -> str:
    """Lấy URL Google Sheet Webhook từ Streamlit secrets, biến môi trường hoặc file config."""
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
    if CONFIG_FILE.is_file():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("webhook_url", "").strip()
        except Exception:
            pass
    return ""

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
    # Tự động lấy thông tin máy nếu không truyền
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if not device_info:
        try:
            hostname = socket.gethostname()
            os_name = f"{platform.system()} {platform.release()}"
            device_info = f"{hostname} ({os_name})"
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
