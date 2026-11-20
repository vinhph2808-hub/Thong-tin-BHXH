# -*- coding: utf-8 -*-
"""
license_dialog.py - Giao diện kích hoạt và quản trị bản quyền Online (Google Sheets).
Đồng bộ và gộp chung quản lý với Tool Thuế.
"""

from __future__ import annotations

import datetime
import re
import tkinter as tk
import webbrowser
from tkinter import messagebox

from .license_manager import (
    check_license_online,
    generate_activation_key,
    get_device_id,
    is_master_device,
    load_installed_license,
    save_installed_license,
    verify_activation_key,
)


class ActivationDialog(tk.Toplevel):
    """Cửa sổ yêu cầu phê duyệt / kích hoạt bản quyền khi chạy trên máy khác."""

    def __init__(self, parent=None, initial_status="CHỜ DUYỆT", initial_message=""):
        super().__init__(parent)
        self.device_id = get_device_id()
        self.activated = False
        self.license_desc = ""
        self.current_status = initial_status
        self.current_message = initial_message or "Thiết bị của bạn đang chờ Quản trị viên phê duyệt trên Google Sheets."

        self.title("Xác thực bản quyền phần mềm - BHXH C12")
        self.geometry("580x560")
        self.resizable(False, False)
        self.configure(bg="#F8FAFC")

        # Căn giữa màn hình
        self.update_idletasks()
        w = 580
        h = 560
        x = max(0, (self.winfo_screenwidth() // 2) - (w // 2))
        y = max(0, (self.winfo_screenheight() // 2) - (h // 2) - 30)
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        if parent:
            self.transient(parent)
        self.grab_set()

        self._build_ui()

    def _build_ui(self):
        # Chọn màu sắc thanh tiêu đề theo trạng thái
        if "KHÓA" in self.current_status or "BLOCKED" in self.current_status:
            hdr_bg = "#991B1B"  # Đỏ đậm
            hdr_title = "THIẾT BỊ ĐÃ BỊ NGẮT KẾT NỐI"
            hdr_sub = "Quản trị viên đã khóa quyền sử dụng phần mềm trên thiết bị này"
            hdr_icon = "🚫"
        elif "LỖI" in self.current_status or "MẠNG" in self.current_status:
            hdr_bg = "#B45309"  # Cam đậm
            hdr_title = "YÊU CẦU KẾT NỐI INTERNET"
            hdr_sub = "Cần kết nối mạng để xác thực quyền sử dụng từ Quản trị viên"
            hdr_icon = "⚠️"
        else:
            hdr_bg = "#1E3A8A"  # Xanh dương đậm
            hdr_title = "HỆ THỐNG KIỂM SOÁT BẢN QUYỀN ONLINE"
            hdr_sub = "Thiết bị mới cần được Quản trị viên duyệt trên Google Sheets"
            hdr_icon = "🌐"

        header = tk.Frame(self, bg=hdr_bg, height=75)
        header.pack(fill="x")
        header.pack_propagate(False)

        h_content = tk.Frame(header, bg=hdr_bg)
        h_content.pack(fill="both", expand=True, padx=20, pady=10)

        lbl_icon = tk.Label(h_content, text=hdr_icon, font=("Segoe UI Emoji", 22), bg=hdr_bg, fg="#FFFFFF")
        lbl_icon.pack(side="left", padx=(0, 12))

        title_box = tk.Frame(h_content, bg=hdr_bg)
        title_box.pack(side="left", fill="both", expand=True)

        tk.Label(title_box, text=hdr_title, font=("Segoe UI", 12, "bold"), fg="#FFFFFF", bg=hdr_bg).pack(anchor="w")
        tk.Label(title_box, text=hdr_sub, font=("Segoe UI", 9), fg="#E2E8F0", bg=hdr_bg).pack(anchor="w", pady=(2, 0))

        # Body
        body = tk.Frame(self, bg="#F8FAFC", padx=24, pady=14)
        body.pack(fill="both", expand=True)

        # Card 1: Thông báo trạng thái hiện tại
        status_card = tk.Frame(body, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, padx=14, pady=10)
        status_card.pack(fill="x", pady=(0, 12))

        tk.Label(status_card, text="TRẠNG THÁI HIỆN TẠI:", font=("Segoe UI", 8, "bold"), fg="#64748B", bg="#FFFFFF").pack(anchor="w")
        self.lbl_msg = tk.Label(
            status_card,
            text=self.current_message,
            font=("Segoe UI", 9, "bold"),
            fg="#1E293B",
            bg="#FFFFFF",
            wraplength=500,
            justify="left",
        )
        self.lbl_msg.pack(anchor="w", pady=(4, 0))

        # Card 2: Mã thiết bị của khách
        card_id = tk.Frame(body, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=14, pady=10)
        card_id.pack(fill="x", pady=(0, 12))

        tk.Label(card_id, text="MÃ THIẾT BỊ CỦA BẠN (HARDWARE ID):", font=("Segoe UI", 8, "bold"), fg="#64748B", bg="#FFFFFF").pack(anchor="w")

        id_row = tk.Frame(card_id, bg="#FFFFFF")
        id_row.pack(fill="x", pady=(6, 8))

        lbl_hwid = tk.Label(id_row, text=self.device_id, font=("Consolas", 15, "bold"), fg="#1E40AF", bg="#EFF6FF", padx=10, pady=4)
        lbl_hwid.pack(side="left")

        btn_copy = tk.Button(
            id_row,
            text="📋 Sao chép mã",
            font=("Segoe UI", 9, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=12,
            pady=5,
            cursor="hand2",
            command=self.copy_device_id,
        )
        btn_copy.pack(side="right")

        # Nút tiện ích gửi Zalo
        btn_zalo_msg = tk.Button(
            card_id,
            text="💬 Sao chép mẫu tin nhắn gửi Quản trị viên qua Zalo để duyệt nhanh",
            font=("Segoe UI", 9),
            bg="#F1F5F9",
            fg="#334155",
            activebackground="#E2E8F0",
            relief="flat",
            padx=8,
            pady=4,
            cursor="hand2",
            command=self.copy_zalo_message,
        )
        btn_zalo_msg.pack(fill="x")

        # Card 3: Nút kiểm tra lại trạng thái duyệt Online (Tự động mở khóa)
        card_recheck = tk.Frame(body, bg="#F0FDF4", highlightbackground="#BBF7D0", highlightthickness=1, padx=14, pady=12)
        card_recheck.pack(fill="x", pady=(0, 12))

        tk.Label(card_recheck, text="BẠN ĐÃ LIÊN HỆ QUẢN TRỊ VIÊN DUYỆT TRÊN GOOGLE SHEETS?", font=("Segoe UI", 8, "bold"), fg="#166534", bg="#F0FDF4").pack(anchor="w")

        self.btn_recheck = tk.Button(
            card_recheck,
            text="🔄 KIỂM TRA LẠI TRẠNG THÁI DUYỆT ONLINE",
            font=("Segoe UI", 10, "bold"),
            bg="#16A34A",
            fg="#FFFFFF",
            activebackground="#15803D",
            activeforeground="#FFFFFF",
            relief="flat",
            pady=8,
            cursor="hand2",
            command=self.recheck_online,
        )
        self.btn_recheck.pack(fill="x", pady=(6, 2))

        # Card 4: Nhập mã kích hoạt offline khẩn cấp
        card_act = tk.Frame(body, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=14, pady=8)
        card_act.pack(fill="x", pady=(0, 10))

        tk.Label(card_act, text="HOẶC NHẬP MÃ KÍCH HOẠT KHẨN CẤP (NẾU CÓ):", font=("Segoe UI", 8, "bold"), fg="#64748B", bg="#FFFFFF").pack(anchor="w")

        act_row = tk.Frame(card_act, bg="#FFFFFF")
        act_row.pack(fill="x", pady=(4, 2))

        self.key_entry = tk.Entry(act_row, font=("Consolas", 10), bg="#F8FAFC", relief="solid", bd=1)
        self.key_entry.pack(side="left", fill="x", expand=True, ipady=4)
        self.key_entry.bind("<Return>", lambda _: self.activate_offline())

        btn_act_manual = tk.Button(
            act_row,
            text="Kích hoạt",
            font=("Segoe UI", 9, "bold"),
            bg="#475569",
            fg="#FFFFFF",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=12,
            pady=3,
            cursor="hand2",
            command=self.activate_offline,
        )
        btn_act_manual.pack(side="right", padx=(6, 0))

        # Footer
        btn_bar = tk.Frame(body, bg="#F8FAFC")
        btn_bar.pack(fill="x", side="bottom")

        btn_exit = tk.Button(
            btn_bar,
            text="Thoát phần mềm",
            font=("Segoe UI", 9),
            bg="#E2E8F0",
            fg="#334155",
            activebackground="#CBD5E1",
            relief="flat",
            padx=16,
            pady=6,
            cursor="hand2",
            command=self.on_cancel,
        )
        btn_exit.pack(side="left")

    def copy_device_id(self):
        self.clipboard_clear()
        self.clipboard_append(self.device_id)
        messagebox.showinfo("Đã sao chép", f"Đã sao chép Mã thiết bị:\n\n{self.device_id}\n\nHãy gửi mã này cho Quản trị viên để được phê duyệt.")

    def copy_zalo_message(self):
        msg = f"Chào bạn, mình vừa mở Tool BHXH C12 trên máy tính. Nhờ bạn mở Google Sheets phê duyệt quyền sử dụng cho máy mình nhé!\nMã thiết bị: {self.device_id}"
        self.clipboard_clear()
        self.clipboard_append(msg)
        messagebox.showinfo("Đã sao chép mẫu tin nhắn", f"Đã sao chép tin nhắn vào Clipboard:\n\n{msg}\n\nBạn có thể dán (Ctrl+V) gửi ngay qua Zalo hoặc Email cho Quản trị viên.")

    def recheck_online(self):
        self.btn_recheck.config(state="disabled", text="⏳ Đang kiểm tra kết nối...")
        self.update()
        try:
            valid, msg, desc = check_license_online(self.device_id)
            if valid:
                self.activated = True
                self.license_desc = desc
                messagebox.showinfo("Phê duyệt thành công", f"Chúc mừng! {msg}\n\nPhần mềm đã sẵn sàng sử dụng.")
                self.destroy()
                return
            else:
                self.current_message = msg
                self.lbl_msg.config(text=msg)
                messagebox.showwarning("Chưa được phê duyệt", msg)
        finally:
            self.btn_recheck.config(state="normal", text="🔄 KIỂM TRA LẠI TRẠNG THÁI DUYỆT ONLINE")

    def activate_offline(self):
        input_key = self.key_entry.get().strip()
        valid, msg, desc = verify_activation_key(self.device_id, input_key)
        if not valid:
            messagebox.showerror("Kích hoạt không thành công", msg)
            return

        save_installed_license(input_key)
        self.activated = True
        self.license_desc = desc
        messagebox.showinfo("Kích hoạt thành công", f"Chúc mừng! {msg}\n\nPhần mềm đã sẵn sàng sử dụng.")
        self.destroy()

    def on_cancel(self):
        self.activated = False
        self.destroy()


class AdminLicenseDialog(tk.Toplevel):
    """Cửa sổ dành cho Quản trị viên (bạn) để quản lý bản quyền & ngắt kết nối từ xa."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.title("Quản lý bản quyền & Phê duyệt thiết bị (Admin)")
        self.geometry("580x520")
        self.resizable(False, False)
        self.configure(bg="#F8FAFC")

        self.update_idletasks()
        w = 580
        h = 520
        x = max(0, (self.winfo_screenwidth() // 2) - (w // 2))
        y = max(0, (self.winfo_screenheight() // 2) - (h // 2) - 30)
        self.geometry(f"{w}x{h}+{x}+{y}")

        if parent:
            self.transient(parent)
        self.grab_set()

        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self, bg="#1E3A8A", height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        h_content = tk.Frame(header, bg="#1E3A8A")
        h_content.pack(fill="both", expand=True, padx=20, pady=10)

        lbl_icon = tk.Label(h_content, text="👑", font=("Segoe UI Emoji", 20), bg="#1E3A8A", fg="#FFFFFF")
        lbl_icon.pack(side="left", padx=(0, 10))

        title_box = tk.Frame(h_content, bg="#1E3A8A")
        title_box.pack(side="left", fill="both", expand=True)

        tk.Label(title_box, text="TRUNG TÂM QUẢN TRỊ BẢN QUYỀN (ADMIN)", font=("Segoe UI", 12, "bold"), fg="#FFFFFF", bg="#1E3A8A").pack(anchor="w")
        tk.Label(title_box, text="Quản lý phê duyệt & ngắt kết nối thiết bị từ xa qua Google Sheets (chung với Tool Thuế)", font=("Segoe UI", 9), fg="#93C5FD", bg="#1E3A8A").pack(anchor="w", pady=(2, 0))

        body = tk.Frame(self, bg="#F8FAFC", padx=24, pady=14)
        body.pack(fill="both", expand=True)

        # PHẦN 1: Quản lý Online qua Google Sheets
        card_online = tk.Frame(body, bg="#FFFFFF", highlightbackground="#93C5FD", highlightthickness=1, padx=16, pady=12)
        card_online.pack(fill="x", pady=(0, 14))

        tk.Label(card_online, text="🌐 QUẢN LÝ ONLINE TỪ XA QUA GOOGLE SHEETS:", font=("Segoe UI", 9, "bold"), fg="#1E40AF", bg="#FFFFFF").pack(anchor="w")
        desc_online = (
            "Bạn có thể mở Google Sheets trên điện thoại hoặc máy tính để duyệt hoặc\n"
            "chủ động NGẮT KẾT NỐI (KHÓA) bất kỳ máy nào chỉ bằng 1 cú click chuột!"
        )
        tk.Label(card_online, text=desc_online, font=("Segoe UI", 9), fg="#475569", bg="#FFFFFF", justify="left").pack(anchor="w", pady=(4, 10))

        btn_row = tk.Frame(card_online, bg="#FFFFFF")
        btn_row.pack(fill="x")

        btn_open_sheet = tk.Button(
            btn_row,
            text="📊 MỞ GOOGLE SHEETS QUẢN LÝ",
            font=("Segoe UI", 10, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=14,
            pady=7,
            cursor="hand2",
            command=self.open_google_sheets,
        )
        btn_open_sheet.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_test_conn = tk.Button(
            btn_row,
            text="🔍 Kiểm tra kết nối",
            font=("Segoe UI", 9),
            bg="#F1F5F9",
            fg="#334155",
            activebackground="#E2E8F0",
            relief="flat",
            padx=10,
            pady=7,
            cursor="hand2",
            command=self.test_connection,
        )
        btn_test_conn.pack(side="right")

        # PHẦN 2: Cấp mã offline dự phòng
        card_offline = tk.Frame(body, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=16, pady=12)
        card_offline.pack(fill="x")

        tk.Label(card_offline, text="🔑 CẤP MÃ KÍCH HOẠT OFFLINE DỰ PHÒNG:", font=("Segoe UI", 9, "bold"), fg="#334155", bg="#FFFFFF").pack(anchor="w")

        id_input_frame = tk.Frame(card_offline, bg="#FFFFFF")
        id_input_frame.pack(fill="x", pady=(6, 8))

        self.client_id_entry = tk.Entry(id_input_frame, font=("Consolas", 11), bg="#F8FAFC", relief="solid", bd=1)
        self.client_id_entry.pack(side="left", fill="x", expand=True, ipady=3)

        btn_paste = tk.Button(
            id_input_frame,
            text="📋 Dán",
            font=("Segoe UI", 9),
            bg="#F1F5F9",
            fg="#334155",
            activebackground="#E2E8F0",
            relief="flat",
            padx=10,
            pady=2,
            cursor="hand2",
            command=self.paste_from_clipboard,
        )
        btn_paste.pack(side="right", padx=(6, 0))

        # Chọn thời hạn
        exp_frame = tk.Frame(card_offline, bg="#FFFFFF")
        exp_frame.pack(fill="x", pady=(0, 8))

        self.expiry_var = tk.StringVar(value="PERM")
        for text, val in [("Vĩnh viễn", "PERM"), ("1 Năm", "1Y"), ("6 Tháng", "6M"), ("1 Tháng", "1M")]:
            tk.Radiobutton(
                exp_frame,
                text=text,
                value=val,
                variable=self.expiry_var,
                bg="#FFFFFF",
                activebackground="#FFFFFF",
                font=("Segoe UI", 9),
            ).pack(side="left", padx=(0, 12))

        btn_generate = tk.Button(
            card_offline,
            text="⚡ TẠO & SAO CHÉP MÃ KÍCH HOẠT OFFLINE",
            font=("Segoe UI", 9, "bold"),
            bg="#475569",
            fg="#FFFFFF",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief="flat",
            pady=6,
            cursor="hand2",
            command=self.do_generate,
        )
        btn_generate.pack(fill="x", pady=(2, 6))

        self.res_entry = tk.Entry(card_offline, font=("Consolas", 10, "bold"), bg="#EFF6FF", fg="#1E40AF", relief="solid", bd=1)
        self.res_entry.pack(fill="x", ipady=3)

        self.lbl_status = tk.Label(card_offline, text="", font=("Segoe UI", 8), fg="#16A34A", bg="#FFFFFF")
        self.lbl_status.pack(anchor="w", pady=(3, 0))

    def open_google_sheets(self):
        webbrowser.open("https://docs.google.com/spreadsheets/u/0/")

    def test_connection(self):
        try:
            valid, msg, desc = check_license_online("TK-TEST-PING-0001")
            messagebox.showinfo(
                "Kiểm tra kết nối thành công",
                f"✔ Kết nối máy chủ Google Sheets hoạt động tốt!\n\nPhản hồi máy chủ: {msg}",
            )
        except Exception as exc:
            messagebox.showerror("Lỗi kết nối", f"Không thể kết nối máy chủ Google Sheets:\n{exc}")

    def paste_from_clipboard(self):
        try:
            text = self.clipboard_get().strip()
            m = re.search(r"TK-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}", text, re.I)
            if m:
                self.client_id_entry.delete(0, tk.END)
                self.client_id_entry.insert(0, m.group(0).upper())
            else:
                self.client_id_entry.delete(0, tk.END)
                self.client_id_entry.insert(0, text)
        except Exception:
            pass

    def do_generate(self):
        client_id = self.client_id_entry.get().strip().upper()
        if not client_id or not client_id.startswith("TK-"):
            messagebox.showerror("Mã thiết bị không hợp lệ", "Mã thiết bị phải có định dạng chuẩn: TK-XXXX-XXXX-XXXX")
            return

        choice = self.expiry_var.get()
        today = datetime.date.today()
        if choice == "PERM":
            exp_code = "PERM"
            exp_desc = "Vĩnh viễn"
        elif choice == "1Y":
            exp_date = today + datetime.timedelta(days=365)
            exp_code = exp_date.strftime("%Y%m%d")
            exp_desc = f"1 Năm (đến {exp_date.strftime('%d/%m/%Y')})"
        elif choice == "6M":
            exp_date = today + datetime.timedelta(days=180)
            exp_code = exp_date.strftime("%Y%m%d")
            exp_desc = f"6 Tháng (đến {exp_date.strftime('%d/%m/%Y')})"
        elif choice == "1M":
            exp_date = today + datetime.timedelta(days=30)
            exp_code = exp_date.strftime("%Y%m%d")
            exp_desc = f"1 Tháng (đến {exp_date.strftime('%d/%m/%Y')})"
        else:
            exp_code = "PERM"
            exp_desc = "Vĩnh viễn"

        key = generate_activation_key(client_id, exp_code)
        self.res_entry.delete(0, tk.END)
        self.res_entry.insert(0, key)

        self.clipboard_clear()
        self.clipboard_append(key)
        self.lbl_status.config(
            text=f"✔ Đã tạo và sao chép mã ({exp_desc})! Bạn có thể dán (Ctrl+V) gửi cho khách.",
            fg="#16A34A",
        )


def ensure_authorized(parent=None) -> bool:
    """Kiểm tra bản quyền trực tuyến khi khởi động:
    - Nếu là Master Admin: Cho qua ngay lập tức.
    - Nếu là Client: Kết nối Google Sheets kiểm tra.
      + Nếu ACTIVE: Cho qua ngay lập tức.
      + Nếu PENDING / BLOCKED / Mất mạng: Mở ActivationDialog.
    """
    dev_id = get_device_id()
    if is_master_device(dev_id):
        return True

    valid, msg, desc = load_installed_license(dev_id)
    if valid:
        return True

    dlg = ActivationDialog(parent, initial_status=desc, initial_message=msg)
    dlg.wait_window()
    return dlg.activated
