# -*- coding: utf-8 -*-
"""
license_manager.py - Hệ thống quản lý bản quyền & xác thực thiết bị (Hardware ID).
Tích hợp kiểm soát Online từ xa qua Google Sheets, đồng bộ chung với Tool Thuế.
Quản trị viên có thể ngắt kết nối (khóa) hoặc phê duyệt bất kỳ máy nào từ xa bất kỳ lúc nào.
"""

from __future__ import annotations

import ctypes
import datetime
import hashlib
import hmac
import json
import os
import platform
import urllib.parse
import urllib.request
import winreg
from pathlib import Path

# URL Google Apps Script Web App kiểm soát bản quyền trực tuyến từ xa (gộp chung với Tool Thuế)
ONLINE_SERVER_URL = "https://script.google.com/macros/s/AKfycbxshsd1_zsRzOBALdfCd9zVh29VlpTsmx-rxCstIbYCmvJbwKuZNKFIoTFfqUb4Cxxj/exec"

# Khóa bí mật ký số HMAC-SHA256 (đồng bộ bảo mật)
_SECRET_SALT = b"Antigravity_Tax_Tool_VinhPH_2026_@_SecureLicensing_Salt_9981"

# Danh sách Mã thiết bị Quản trị (Master Admin - Máy của bạn, tự động toàn quyền)
MASTER_DEVICE_IDS = {
    "TK-E790-3578-652A",  # Laptop hiện tại của bạn
}
_MASTER_COMPUTER_NAMES = {"SG-KI-VINHPH-UI"}
_MASTER_USERNAMES = {"VINHPH"}

# Khóa vạn năng dành riêng cho Quản trị viên (dùng để kích hoạt khẩn cấp trên bất kỳ máy nào)
MASTER_UNLOCK_KEY = "ACT-MASTER-ADMIN-VINHPH-2026"


def get_device_id() -> str:
    """Tạo mã định danh phần cứng duy nhất (Hardware ID) cho máy tính Windows.
    Đồng bộ tuyệt đối 100% với thuật toán của Tool Thuế.
    """
    components: list[str] = []

    # 1. MachineGuid từ Windows Registry
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
            guid, _ = winreg.QueryValueEx(key, "MachineGuid")
            components.append(str(guid).strip())
    except Exception:
        pass

    # 2. Volume Serial Number của ổ C:
    try:
        vol_serial = ctypes.c_ulong()
        ctypes.windll.kernel32.GetVolumeInformationW(
            ctypes.c_wchar_p("C:\\"),
            None, 0,
            ctypes.byref(vol_serial),
            None, None, None, 0,
        )
        components.append(hex(vol_serial.value)[2:].upper())
    except Exception:
        pass

    # 3. Processor Identifier từ biến môi trường
    proc_id = os.environ.get("PROCESSOR_IDENTIFIER", "").strip()
    if proc_id:
        components.append(proc_id)

    # Dự phòng nếu không lấy được thông tin hệ thống
    if not components:
        components.append(platform.node())
        components.append(os.environ.get("USERNAME", "DEFAULT_USER"))

    raw = "|".join(components)
    hwid = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:12].upper()
    return f"TK-{hwid[:4]}-{hwid[4:8]}-{hwid[8:12]}"


def is_master_device(device_id: str | None = None) -> bool:
    """Kiểm tra xem thiết bị hiện tại có phải là máy Quản trị viên (bạn) hay không."""
    if device_id is None:
        device_id = get_device_id()
    if device_id.strip().upper() in MASTER_DEVICE_IDS:
        return True
    
    curr_comp = (os.environ.get("COMPUTERNAME") or "").strip().upper()
    curr_user = (os.environ.get("USERNAME") or "").strip().upper()
    if curr_comp in _MASTER_COMPUTER_NAMES and curr_user in _MASTER_USERNAMES:
        return True
    return False


def is_master_machine() -> bool:
    """Tương thích ngược với các module gọi is_master_machine()."""
    return is_master_device()


def get_license_file_path() -> Path:
    """Đường dẫn tệp lưu bản quyền trên máy client."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base_dir = Path(local_app_data) / "C12_BHXH_Tool"
    else:
        base_dir = Path.home() / ".c12_bhxh_tool"
    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir / "license.lic"


def save_installed_license(key: str) -> None:
    """Lưu mã kích hoạt vào tệp bản quyền."""
    lic_path = get_license_file_path()
    lic_path.write_text(key.strip().upper(), encoding="utf-8")


def generate_activation_key(device_id: str, expiry: str = "PERM") -> str:
    """Tạo mã kích hoạt bản quyền cho một mã thiết bị cụ thể."""
    device_id = device_id.strip().upper()
    expiry = expiry.strip().upper()
    msg = f"{device_id}:{expiry}".encode("utf-8")
    sig = hmac.new(_SECRET_SALT, msg, hashlib.sha256).hexdigest()[:12].upper()
    return f"ACT-{expiry}-{sig[:4]}-{sig[4:8]}-{sig[8:12]}"


def verify_activation_key(device_id: str, key: str) -> tuple[bool, str, str]:
    """Xác thực mã kích hoạt offline cho thiết bị.
    Trả về (hợp_lệ: bool, thông_báo: str, loại_hạn: str).
    """
    key = key.strip().upper()
    device_id = device_id.strip().upper()

    if not key:
        return False, "Vui lòng nhập mã kích hoạt.", ""

    # Kiểm tra khóa vạn năng Master Admin
    if key == MASTER_UNLOCK_KEY:
        return True, "Kích hoạt thành công bằng Mã Quản trị đặc biệt.", "Quản trị viên (Vĩnh viễn)"

    parts = key.split("-")
    if len(parts) != 5 or parts[0] != "ACT":
        return False, "Mã kích hoạt không đúng định dạng chuẩn (ACT-...).", ""

    expiry = parts[1]
    sig_input = "".join(parts[2:])
    expected_sig = hmac.new(_SECRET_SALT, f"{device_id}:{expiry}".encode("utf-8"), hashlib.sha256).hexdigest()[:12].upper()

    if not hmac.compare_digest(sig_input, expected_sig):
        return False, "Mã kích hoạt không hợp lệ cho thiết bị này.", ""

    if expiry == "PERM":
        return True, "Kích hoạt bản quyền vĩnh viễn thành công.", "Bản quyền vĩnh viễn"

    # Kiểm tra hạn dùng
    try:
        exp_date = datetime.datetime.strptime(expiry, "%Y%m%d").date()
        today = datetime.date.today()
        if today > exp_date:
            return False, f"Mã kích hoạt đã hết hạn vào ngày {exp_date.strftime('%d/%m/%Y')}.", ""
        days_left = (exp_date - today).days
        return True, f"Kích hoạt thành công. Hạn dùng đến {exp_date.strftime('%d/%m/%Y')} (còn {days_left} ngày).", f"Đến {exp_date.strftime('%d/%m/%Y')}"
    except ValueError:
        return False, "Thời hạn trong mã kích hoạt không hợp lệ.", ""


def check_license_online(device_id: str | None = None) -> tuple[bool, str, str]:
    """Kiểm tra bản quyền trực tuyến qua máy chủ Google Sheets.
    Trả về: (cho_phép_chạy: bool, thông_báo: str, trạng_thái_mô_tả: str)
    """
    if device_id is None:
        device_id = get_device_id()

    # Máy Master Admin của bạn luôn luôn được phép toàn quyền (dù có mạng hay mất mạng)
    if is_master_device(device_id):
        return True, "Thiết bị Quản trị viên (Master Admin)", "Quản trị viên (Toàn quyền)"

    server_url = ONLINE_SERVER_URL.strip()
    if not server_url:
        return False, "Chưa cấu hình máy chủ bản quyền trực tuyến.", "CHƯA CẤU HÌNH"

    params = {
        "action": "check",
        "device_id": device_id,
        "hostname": f"[BHXH] {platform.node()}",
        "username": os.environ.get("USERNAME", "USER"),
        "tool": "BHXH_C12",
    }
    full_url = f"{server_url}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(full_url, headers={"User-Agent": "AntigravityBHXHTool/2026"})

    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            content = resp.read().decode("utf-8")
            data = json.loads(content)
            status = str(data.get("status", "")).upper()
            msg = str(data.get("message", ""))
            expiry = str(data.get("expiry", "Vĩnh viễn"))

            if status == "ACTIVE":
                return True, msg or "Thiết bị đã được phê duyệt hợp lệ.", f"Đã duyệt ({expiry})"
            elif status == "BLOCKED":
                # Xóa file bản quyền offline nếu có để vô hiệu hóa hoàn toàn ngay lập tức
                try:
                    lic_p = get_license_file_path()
                    if lic_p.is_file():
                        lic_p.unlink()
                except Exception:
                    pass
                return False, msg or "Thiết bị này đã bị Quản trị viên ngắt kết nối và khóa quyền sử dụng.", "BỊ KHÓA TỪ XA"
            elif status == "EXPIRED":
                return False, msg or "Bản quyền trên thiết bị này đã hết hạn.", "HẾT HẠN SỬ DỤNG"
            elif status == "PENDING":
                return False, msg or "Thiết bị đang chờ Quản trị viên phê duyệt trên Google Sheets.", "CHỜ DUYỆT"
            else:
                return False, msg or "Trạng thái thiết bị không hợp lệ.", status
    except Exception:
        # Nếu mất kết nối Internet, kiểm tra giấy phép offline dự phòng
        lic_file = get_license_file_path()
        if lic_file.is_file():
            saved_key = lic_file.read_text(encoding="utf-8").strip()
            val, off_msg, off_desc = verify_activation_key(device_id, saved_key)
            if val:
                return True, "Chạy chế độ Offline (Đã xác thực trước đó)", off_desc

        return False, "Không thể kết nối máy chủ quản lý bản quyền Google Sheets. Vui lòng kết nối Internet để xác thực quyền sử dụng từ Quản trị viên.", "LỖI KẾT NỐI MẠNG"


def load_installed_license(device_id: str | None = None) -> tuple[bool, str, str]:
    """Kiểm tra bản quyền cho máy tính."""
    if device_id is None:
        device_id = get_device_id()

    # Master Admin luôn luôn hợp lệ
    if is_master_device(device_id):
        return True, "Thiết bị Quản trị viên (Master Admin)", "Quản trị viên (Toàn quyền)"

    # Kiểm tra Online là ưu tiên số 1
    return check_license_online(device_id)


def check_activation_status() -> tuple[bool, str]:
    """Hàm tương thích ngược trả về (bool, str)."""
    ok, msg, desc = load_installed_license()
    return ok, desc or msg
