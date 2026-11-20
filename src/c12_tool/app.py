from __future__ import annotations

import os
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import (
    BOTH,
    BOTTOM,
    END,
    LEFT,
    RIGHT,
    TOP,
    X,
    Y,
    BooleanVar,
    Canvas,
    Frame,
    Label,
    Menu,
    StringVar,
    Toplevel,
    filedialog,
    messagebox,
)
from tkinter import ttk

from tkinterdnd2 import DND_FILES, TkinterDnD

from .excel_writer import export_records
from .license_dialog import AdminLicenseDialog, ensure_authorized
from .license_manager import is_master_machine
from .parser import C12ParseError, C12Record, parse_c12_records

APP_TITLE = "Xuất thông báo BHXH C12"
APP_VERSION = "v2.0 Pro"


SUPPORTED_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"}


def enable_high_dpi_awareness() -> None:
    """Enable high-DPI scaling on Windows for crisp modern fonts."""
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def app_folder() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def bundled_template() -> Path | None:
    """Return the workbook embedded by PyInstaller (or the source copy in development)."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        candidate = base / "assets" / "THÔNG BÁO BXH.xlsx"
    else:
        candidate = app_folder() / "THÔNG BÁO BXH.xlsx"
    return candidate if candidate.is_file() else None


def find_default_template() -> Path | None:
    # An external workbook can override the built-in one during local development.
    candidates = list(app_folder().glob("*.xlsx")) + list(Path.cwd().glob("*.xlsx"))
    for path in candidates:
        upper = path.name.upper()
        if "THÔNG BÁO" in upper or "THONG BAO" in upper or "BXH" in upper or "BHXH" in upper:
            if "KET_QUA" not in upper:
                return path
    return bundled_template()


def format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"


def detect_unit_hint(filename: str) -> str:
    upper = filename.upper()
    if "YN" in upper:
        return "YN"
    elif "BW" in upper:
        return "IC (BW)"
    elif "IC" in upper:
        return "IC"
    return "Tự động"


class ModernButton(Label):
    """Custom flat button with smooth hover effects and rounded aesthetic."""

    def __init__(
        self,
        master,
        text: str,
        command=None,
        bg: str = "#2563EB",
        fg: str = "#FFFFFF",
        hover_bg: str = "#1D4ED8",
        active_bg: str = "#1E40AF",
        font=("Segoe UI", 10, "bold"),
        padx: int = 14,
        pady: int = 7,
        cursor: str = "hand2",
        **kwargs,
    ):
        super().__init__(
            master,
            text=text,
            bg=bg,
            fg=fg,
            font=font,
            padx=padx,
            pady=pady,
            cursor=cursor,
            relief="flat",
            **kwargs,
        )
        self.command = command
        self.normal_bg = bg
        self.hover_bg = hover_bg
        self.active_bg = active_bg
        self.normal_fg = fg
        self._disabled = False

        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _on_enter(self, _event) -> None:
        if not self._disabled:
            self.configure(bg=self.hover_bg)

    def _on_leave(self, _event) -> None:
        if not self._disabled:
            self.configure(bg=self.normal_bg)

    def _on_press(self, _event) -> None:
        if not self._disabled:
            self.configure(bg=self.active_bg)

    def _on_release(self, event) -> None:
        if not self._disabled:
            self.configure(bg=self.hover_bg)
            if self.command:
                self.command()

    def set_state(self, state: str) -> None:
        if state == "disabled":
            self._disabled = True
            self.configure(bg="#CBD5E1", fg="#94A3B8", cursor="arrow")
        else:
            self._disabled = False
            self.configure(bg=self.normal_bg, fg=self.normal_fg, cursor="hand2")


class C12App(TkinterDnD.Tk):
    def __init__(self, unrestricted: bool = False) -> None:
        self.is_unrestricted = unrestricted
        enable_high_dpi_awareness()
        super().__init__()
        self.title(f"{APP_TITLE} - {APP_VERSION}")
        self.geometry("980x700")
        self.minsize(880, 580)
        self.configure(bg="#F1F5F9")

        self.pdfs: list[Path] = []
        self.template: Path | None = find_default_template()
        self.output_dir = app_folder() / "Ket_qua"
        self.auto_open_var = BooleanVar(value=True)

        self._center_window()
        self._configure_styles()
        self._build_ui()
        self._update_stats()

    def _center_window(self) -> None:
        self.update_idletasks()
        width = 980
        height = 700
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _configure_styles(self) -> None:
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        # Global Treeview style
        self.style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground="#1E293B",
            rowheight=26,
            fieldbackground="#FFFFFF",
            font=("Segoe UI", 9),
            borderwidth=0,
        )
        self.style.configure(
            "Treeview.Heading",
            background="#F8FAFC",
            foreground="#334155",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padding=(6, 4),
        )
        self.style.map(
            "Treeview.Heading",
            background=[("active", "#E2E8F0")],
            foreground=[("active", "#0F172A")],
        )
        self.style.map(
            "Treeview",
            background=[("selected", "#DBEAFE")],
            foreground=[("selected", "#1E40AF")],
        )

        # Progressbar style
        self.style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor="#E2E8F0",
            background="#2563EB",
            thickness=8,
        )

    def _build_ui(self) -> None:
        # 1. Top Header Banner
        header = Frame(self, bg="#0F172A", height=76)
        header.pack(fill=X, side=TOP)
        header.pack_propagate(False)

        header_content = Frame(header, bg="#0F172A")
        header_content.pack(fill=BOTH, expand=True, padx=22, pady=10)

        # Title & Subtitle
        title_box = Frame(header_content, bg="#0F172A")
        title_box.pack(side=LEFT, fill=Y)

        title_row = Frame(title_box, bg="#0F172A")
        title_row.pack(anchor="w")

        title_lbl = Label(
            title_row,
            text=APP_TITLE.upper(),
            font=("Segoe UI", 15, "bold"),
            fg="#FFFFFF",
            bg="#0F172A",
        )
        title_lbl.pack(side=LEFT)

        badge_lbl = Label(
            title_row,
            text=f" {APP_VERSION} ",
            font=("Segoe UI", 8, "bold"),
            fg="#93C5FD",
            bg="#1E293B",
            relief="solid",
            borderwidth=1,
        )
        badge_lbl.pack(side=LEFT, padx=(10, 0))

        sub_lbl = Label(
            title_box,
            text="Trích xuất tự động thông báo C12-TS (YN/IC/BW) sang bảng tính Excel đối chiếu chuyên nghiệp.",
            font=("Segoe UI", 9),
            fg="#94A3B8",
            bg="#0F172A",
        )
        sub_lbl.pack(anchor="w", pady=(3, 0))

        # Right Header Badge (File Counter) & Admin License Button
        self.header_badge = Label(
            header_content,
            text="0 file PDF",
            font=("Segoe UI", 9, "bold"),
            fg="#60A5FA",
            bg="#1E293B",
            padx=12,
            pady=4,
            relief="solid",
            borderwidth=1,
        )
        self.header_badge.pack(side=RIGHT, anchor="center")

        if getattr(self, "is_unrestricted", False):
            # Bản sử dụng tự do: giao diện sạch sẽ, không hiển thị huy hiệu bản quyền
            pass
        elif is_master_machine():
            admin_btn = ModernButton(
                header_content,
                text="🌐 Quản lý Online (Google Sheets)",
                command=self._open_admin_license,
                bg="#2563EB",
                hover_bg="#1D4ED8",
                active_bg="#1E40AF",
                font=("Segoe UI", 9, "bold"),
                padx=10,
                pady=4,
            )
            admin_btn.pack(side=RIGHT, padx=(0, 10), anchor="center")
        else:
            client_badge = Label(
                header_content,
                text="🛡️ Đã kết nối Online",
                font=("Segoe UI", 9, "bold"),
                fg="#4ADE80",
                bg="#1E293B",
                padx=10,
                pady=4,
                relief="solid",
                borderwidth=1,
            )
            client_badge.pack(side=RIGHT, padx=(0, 10), anchor="center")

        # 2. Main Content Container
        main_frame = Frame(self, bg="#F1F5F9", padx=20, pady=14)
        main_frame.pack(fill=BOTH, expand=True)

        # 3. Card 1: File List & Action Area
        card_files = Frame(main_frame, bg="#FFFFFF", relief="solid", borderwidth=1, padx=14, pady=12)
        card_files.configure(highlightbackground="#CBD5E1", highlightcolor="#CBD5E1")
        card_files.pack(fill=BOTH, expand=True)

        # File List Toolbar
        toolbar = Frame(card_files, bg="#FFFFFF")
        toolbar.pack(fill=X, pady=(0, 10))

        lbl_section = Label(
            toolbar,
            text="📁 Danh sách file PDF thông báo C12:",
            font=("Segoe UI", 10, "bold"),
            fg="#0F172A",
            bg="#FFFFFF",
        )
        lbl_section.pack(side=LEFT, padx=(0, 12))

        # Buttons
        btn_add_files = ModernButton(
            toolbar,
            text="➕ Thêm File...",
            command=self.choose_pdfs,
            bg="#2563EB",
            hover_bg="#1D4ED8",
            active_bg="#1E40AF",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
        )
        btn_add_files.pack(side=LEFT, padx=3)

        btn_add_dir = ModernButton(
            toolbar,
            text="📂 Thêm cả thư mục...",
            command=self.choose_directory,
            bg="#F1F5F9",
            fg="#334155",
            hover_bg="#E2E8F0",
            active_bg="#CBD5E1",
            font=("Segoe UI", 9),
            padx=10,
            pady=4,
        )
        btn_add_dir.pack(side=LEFT, padx=3)

        btn_remove = ModernButton(
            toolbar,
            text="🗑️ Xóa file chọn",
            command=self.remove_selected,
            bg="#F1F5F9",
            fg="#334155",
            hover_bg="#FEE2E2",
            active_bg="#FECACA",
            font=("Segoe UI", 9),
            padx=10,
            pady=4,
        )
        btn_remove.pack(side=LEFT, padx=3)

        btn_clear = ModernButton(
            toolbar,
            text="🧹 Xóa tất cả",
            command=self.clear_pdfs,
            bg="#F1F5F9",
            fg="#64748B",
            hover_bg="#E2E8F0",
            active_bg="#CBD5E1",
            font=("Segoe UI", 9),
            padx=10,
            pady=4,
        )
        btn_clear.pack(side=LEFT, padx=3)

        lbl_hint = Label(
            toolbar,
            text="💡 Kéo & thả file PDF / hình ảnh vào bảng",
            font=("Segoe UI", 9, "italic"),
            fg="#64748B",
            bg="#FFFFFF",
        )
        lbl_hint.pack(side=RIGHT)

        # File List Treeview Container
        tree_container = Frame(card_files, bg="#FFFFFF")
        tree_container.pack(fill=BOTH, expand=True)

        columns = ("stt", "name", "unit", "size", "path")
        self.tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            selectmode="extended",
            style="Treeview",
        )

        self.tree.heading("stt", text="#")
        self.tree.heading("name", text="Tên file tài liệu")
        self.tree.heading("unit", text="Đơn vị dự kiến")
        self.tree.heading("size", text="Kích thước")
        self.tree.heading("path", text="Đường dẫn nguồn")

        self.tree.column("stt", width=42, minwidth=35, anchor="center")
        self.tree.column("name", width=310, minwidth=200, anchor="w")
        self.tree.column("unit", width=110, minwidth=80, anchor="center")
        self.tree.column("size", width=95, minwidth=70, anchor="e")
        self.tree.column("path", width=340, minwidth=200, anchor="w")

        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd", background="#F8FAFC")

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # Keyboard shortcuts
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # Enable Drag and Drop
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self.on_drop)

        # 4. Card 2: Options & Output settings
        card_options = Frame(main_frame, bg="#FFFFFF", relief="solid", borderwidth=1, padx=14, pady=10)
        card_options.configure(highlightbackground="#CBD5E1", highlightcolor="#CBD5E1")
        card_options.pack(fill=X, pady=(12, 0))

        # Row 1: Template info
        row_opt1 = Frame(card_options, bg="#FFFFFF")
        row_opt1.pack(fill=X, pady=2)

        lbl_tpl_title = Label(
            row_opt1,
            text="Mẫu Excel:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#FFFFFF",
            width=12,
            anchor="w",
        )
        lbl_tpl_title.pack(side=LEFT)

        self.lbl_tpl_val = Label(
            row_opt1,
            text="Tích hợp sẵn (THÔNG BÁO BXH.xlsx)",
            font=("Segoe UI", 9),
            fg="#0F172A",
            bg="#F1F5F9",
            padx=8,
            pady=2,
            relief="solid",
            borderwidth=1,
        )
        self.lbl_tpl_val.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))

        btn_change_tpl = ModernButton(
            row_opt1,
            text="Đổi mẫu khác...",
            command=self.choose_template,
            bg="#F1F5F9",
            fg="#334155",
            hover_bg="#E2E8F0",
            active_bg="#CBD5E1",
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
        )
        btn_change_tpl.pack(side=LEFT, padx=(0, 4))

        btn_reset_tpl = ModernButton(
            row_opt1,
            text="Mặc định",
            command=self.reset_template,
            bg="#F1F5F9",
            fg="#64748B",
            hover_bg="#E2E8F0",
            active_bg="#CBD5E1",
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
        )
        btn_reset_tpl.pack(side=LEFT)

        # Row 2: Output destination & Checkbox
        row_opt2 = Frame(card_options, bg="#FFFFFF")
        row_opt2.pack(fill=X, pady=(6, 2))

        lbl_out_title = Label(
            row_opt2,
            text="Thư mục kết quả:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#FFFFFF",
            width=12,
            anchor="w",
        )
        lbl_out_title.pack(side=LEFT)

        self.lbl_out_val = Label(
            row_opt2,
            text=str(self.output_dir),
            font=("Segoe UI", 9),
            fg="#0F172A",
            bg="#F1F5F9",
            padx=8,
            pady=2,
            relief="solid",
            borderwidth=1,
            anchor="w",
        )
        self.lbl_out_val.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))

        btn_open_folder = ModernButton(
            row_opt2,
            text="📂 Mở thư mục",
            command=self.open_output_folder,
            bg="#F1F5F9",
            fg="#334155",
            hover_bg="#E2E8F0",
            active_bg="#CBD5E1",
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
        )
        btn_open_folder.pack(side=LEFT, padx=(0, 10))

        chk_auto_open = ttk.Checkbutton(
            row_opt2,
            text="Tự động mở file Excel sau khi xuất",
            variable=self.auto_open_var,
        )
        chk_auto_open.pack(side=RIGHT)

        # 5. Bottom Action & Execution Bar
        bottom_bar = Frame(main_frame, bg="#F1F5F9")
        bottom_bar.pack(fill=X, pady=(12, 0))

        # Left: Status indicator
        status_box = Frame(bottom_bar, bg="#F1F5F9")
        status_box.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 14))

        status_header = Frame(status_box, bg="#F1F5F9")
        status_header.pack(fill=X, anchor="w")

        self.status_icon = Label(status_header, text="⚡", font=("Segoe UI", 10), fg="#2563EB", bg="#F1F5F9")
        self.status_icon.pack(side=LEFT)

        self.status_lbl = Label(
            status_header,
            text="Sẵn sàng xử lý.",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#F1F5F9",
        )
        self.status_lbl.pack(side=LEFT, padx=(4, 0))

        self.progress_lbl = Label(
            status_header,
            text="0%",
            font=("Segoe UI", 9),
            fg="#64748B",
            bg="#F1F5F9",
        )
        self.progress_lbl.pack(side=RIGHT)

        self.progress = ttk.Progressbar(
            status_box,
            mode="determinate",
            style="Accent.Horizontal.TProgressbar",
        )
        self.progress.pack(fill=X, pady=(4, 0))

        # Right: Big Primary Action Button
        self.run_button = ModernButton(
            bottom_bar,
            text="🚀 ĐỌC PDF VÀ XUẤT EXCEL",
            command=self.start_process,
            bg="#2563EB",
            hover_bg="#1D4ED8",
            active_bg="#1E40AF",
            font=("Segoe UI", 11, "bold"),
            padx=24,
            pady=10,
        )
        self.run_button.pack(side=RIGHT)

    def _update_stats(self) -> None:
        count = len(self.pdfs)
        self.header_badge.configure(text=f"{count} file PDF")
        if count == 0:
            self.header_badge.configure(bg="#1E293B", fg="#94A3B8")
            self.status_lbl.configure(text="Vui lòng chọn hoặc kéo thả file PDF hoặc hình ảnh C12 để bắt đầu.")
            self.status_icon.configure(text="⚡", fg="#64748B")
        else:
            self.header_badge.configure(bg="#1E3A8A", fg="#60A5FA")
            self.status_lbl.configure(text=f"Đã nạp {count} file PDF. Sẵn sàng xuất Excel.")
            self.status_icon.configure(text="✓", fg="#10B981")

    def _refresh_list(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, path in enumerate(self.pdfs, 1):
            tag = "even" if i % 2 == 0 else "odd"
            unit_hint = detect_unit_hint(path.name)
            try:
                size_str = format_file_size(path.stat().st_size)
            except Exception:
                size_str = "-"

            self.tree.insert(
                "",
                END,
                values=(i, path.name, unit_hint, size_str, str(path)),
                tags=(tag,),
            )
        self._update_stats()

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file THÔNG BÁO BHXH mẫu",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if path:
            self.template = Path(path)
            self.lbl_tpl_val.configure(text=f"Tùy chỉnh: {self.template.name}")

    def reset_template(self) -> None:
        self.template = find_default_template()
        self.lbl_tpl_val.configure(text="Tích hợp sẵn (THÔNG BÁO BXH.xlsx)")

    def open_output_folder(self) -> None:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(str(self.output_dir))
        else:
            subprocess.Popen(["xdg-open", str(self.output_dir)])

    def choose_pdfs(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Chọn các file PDF hoặc hình ảnh C12",
            filetypes=[
                (
                    "Tài liệu C12 (*.pdf, *.png, *.jpg...)",
                    "*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.webp;*.PDF;*.PNG;*.JPG;*.JPEG;*.BMP;*.TIFF;*.WEBP",
                ),
                ("Tập tin PDF (*.pdf)", "*.pdf;*.PDF"),
                ("Hình ảnh (*.png, *.jpg...)", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.webp;*.PNG;*.JPG;*.JPEG"),
                ("Tất cả tập tin (*.*)", "*.*"),
            ],
        )
        if paths:
            self.add_pdfs([Path(p) for p in paths])

    def choose_directory(self) -> None:
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file PDF / hình ảnh C12")
        if folder:
            dir_path = Path(folder)
            found = [
                p for p in dir_path.iterdir()
                if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS
            ]
            if not found:
                messagebox.showinfo(APP_TITLE, f"Không tìm thấy file PDF hoặc hình ảnh nào trong thư mục:\n{dir_path}")
                return
            self.add_pdfs(found)

    def on_drop(self, event) -> None:
        paths = [Path(p) for p in self.tk.splitlist(event.data)]
        self.add_pdfs(paths)

    def add_pdfs(self, paths: list[Path]) -> None:
        existing = {str(path.resolve()).lower() for path in self.pdfs}
        added = 0
        ignored = 0
        for path in paths:
            if path.is_dir():
                for sub in path.rglob("*"):
                    if sub.is_file() and sub.suffix.lower() in SUPPORTED_EXTS:
                        key = str(sub.resolve()).lower()
                        if key not in existing:
                            self.pdfs.append(sub)
                            existing.add(key)
                            added += 1
                continue
            if not path.is_file() or path.suffix.lower() not in SUPPORTED_EXTS:
                ignored += 1
                continue
            key = str(path.resolve()).lower()
            if key not in existing:
                self.pdfs.append(path)
                existing.add(key)
                added += 1

        self._refresh_list()
        if ignored > 0:
            messagebox.showwarning(APP_TITLE, f"Đã bỏ qua {ignored} file không đúng định dạng hỗ trợ (.pdf, .png, .jpg...).")

    def remove_selected(self) -> None:
        selected_items = self.tree.selection()
        if not selected_items:
            return
        selected_paths = {self.tree.item(item, "values")[4] for item in selected_items}
        self.pdfs = [p for p in self.pdfs if str(p) not in selected_paths]
        self._refresh_list()

    def clear_pdfs(self) -> None:
        self.pdfs.clear()
        self._refresh_list()

    def start_process(self) -> None:
        if not self.template or not self.template.is_file():
            messagebox.showerror(
                APP_TITLE,
                "Không tìm thấy mẫu Excel tích hợp. Vui lòng kiểm tra lại file mẫu THÔNG BÁO BXH.xlsx.",
            )
            return
        if not self.pdfs:
            messagebox.showerror(
                APP_TITLE,
                "Chưa có file nào. Vui lòng bấm 'Thêm file...' hoặc kéo thả file vào bảng.",
            )
            return

        self.run_button.set_state("disabled")
        self.progress.configure(maximum=len(self.pdfs) + 1, value=0)
        self.progress_lbl.configure(text="0%")
        threading.Thread(target=self._process, daemon=True).start()

    def _set_status(self, text: str, progress: int | None = None, icon: str = "⏳", color: str = "#2563EB") -> None:
        def update():
            self.status_lbl.configure(text=text)
            self.status_icon.configure(text=icon, fg=color)
            if progress is not None and self.pdfs:
                self.progress.configure(value=progress)
                pct = int((progress / (len(self.pdfs) + 1)) * 100)
                self.progress_lbl.configure(text=f"{pct}%")

        self.after(0, update)

    def _process(self) -> None:
        records: list[C12Record] = []
        errors: list[str] = []
        duplicates: dict[tuple[str, int], str] = {}
        total_files = len(self.pdfs)

        try:
            for index, path in enumerate(self.pdfs, 1):
                self._set_status(
                    f"Đang đọc ({index}/{total_files}): {path.name}...",
                    index - 1,
                    icon="⏳",
                    color="#2563EB",
                )
                try:
                    parsed_records = parse_c12_records(path)
                    for record in parsed_records:
                        key = (record.unit, record.month)
                        if key in duplicates:
                            errors.append(
                                f"{path.name}: trùng {record.unit} tháng {record.month}; ghi đè bằng dữ liệu file sau cùng."
                            )
                            records = [item for item in records if (item.unit, item.month) != key]
                        duplicates[key] = path.name
                        records.append(record)
                except C12ParseError as exc:
                    errors.append(f"{path.name}: {exc}")
                self._set_status(f"Đã đọc xong {index}/{total_files} PDF", index, icon="⏳", color="#2563EB")

            if not records:
                raise ValueError("Không có file PDF nào đọc thành công dữ liệu C12.\n\n" + "\n".join(errors))

            self._set_status("Đang tạo bảng tính Excel kết quả chuyên nghiệp...", total_files, icon="📊", color="#2563EB")
            output = export_records(self.template, records, self.output_dir)

            self._set_status("Hoàn tất xuất sắc!", total_files + 1, icon="✓", color="#10B981")
            self.after(0, lambda: self.progress_lbl.configure(text="100%"))

            summary = (
                f"Đã trích xuất thành công {len(records)} tháng dữ liệu từ {total_files} file PDF!\n\n"
                f"File kết quả:\n{output.name}\n\n"
                f"Thư mục lưu:\n{output.parent}"
            )
            if errors:
                summary += "\n\nCảnh báo / Ghi chú:\n" + "\n".join(errors[:8])
                if len(errors) > 8:
                    summary += f"\n... và {len(errors) - 8} cảnh báo khác."

            self.after(0, lambda: self._done(output, summary))
        except Exception as exc:
            detail = str(exc) or traceback.format_exc()
            self._set_status("Có lỗi xảy ra trong quá trình xử lý.", icon="⚠️", color="#EF4444")
            self.after(0, lambda: messagebox.showerror(APP_TITLE, detail))
        finally:
            self.after(0, lambda: self.run_button.set_state("normal"))

    def _done(self, output: Path, summary: str) -> None:
        if self.auto_open_var.get():
            try:
                if sys.platform == "win32":
                    os.startfile(str(output))
                else:
                    subprocess.Popen(["xdg-open", str(output)])
            except Exception:
                pass

        # Professional Completion Dialog
        dialog = Toplevel(self)
        dialog.title("Xuất Excel thành công")
        dialog.geometry("540x360")
        dialog.minsize(500, 320)
        dialog.configure(bg="#FFFFFF")
        dialog.transient(self)
        dialog.grab_set()

        # Dialog Header
        dlg_header = Frame(dialog, bg="#10B981", height=50)
        dlg_header.pack(fill=X)
        dlg_header.pack_propagate(False)

        Label(
            dlg_header,
            text="✓  XUẤT BẢO HIỂM XÃ HỘI C12 HOÀN TẤT",
            font=("Segoe UI", 11, "bold"),
            fg="#FFFFFF",
            bg="#10B981",
        ).pack(side=LEFT, padx=16, pady=12)

        # Dialog Body
        dlg_body = Frame(dialog, bg="#FFFFFF", padx=20, pady=16)
        dlg_body.pack(fill=BOTH, expand=True)

        msg_lbl = Label(
            dlg_body,
            text=summary,
            font=("Segoe UI", 9),
            fg="#1E293B",
            bg="#FFFFFF",
            justify=LEFT,
            wraplength=480,
        )
        msg_lbl.pack(anchor="w", fill=BOTH, expand=True)

        # Dialog Actions
        dlg_actions = Frame(dialog, bg="#F8FAFC", padx=16, pady=12)
        dlg_actions.pack(fill=X, side=BOTTOM)

        def open_file():
            try:
                if sys.platform == "win32":
                    os.startfile(str(output))
                else:
                    subprocess.Popen(["xdg-open", str(output)])
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Không thể mở file: {e}")

        def open_folder():
            try:
                if sys.platform == "win32":
                    subprocess.Popen(["explorer", "/select,", str(output)])
                else:
                    subprocess.Popen(["xdg-open", str(output.parent)])
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Không thể mở thư mục: {e}")

        ModernButton(
            dlg_actions,
            text="📊 Mở file Excel",
            command=open_file,
            bg="#2563EB",
            hover_bg="#1D4ED8",
            active_bg="#1E40AF",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
        ).pack(side=LEFT, padx=(0, 8))

        ModernButton(
            dlg_actions,
            text="📂 Mở thư mục chứa file",
            command=open_folder,
            bg="#F1F5F9",
            fg="#334155",
            hover_bg="#E2E8F0",
            active_bg="#CBD5E1",
            font=("Segoe UI", 9),
            padx=12,
            pady=6,
        ).pack(side=LEFT)

        ModernButton(
            dlg_actions,
            text="Đóng",
            command=dialog.destroy,
            bg="#F1F5F9",
            fg="#64748B",
            hover_bg="#E2E8F0",
            active_bg="#CBD5E1",
            font=("Segoe UI", 9),
            padx=14,
            pady=6,
        ).pack(side=RIGHT)

    def _open_admin_license(self) -> None:
        AdminLicenseDialog(self)


def main() -> None:
    if "--self-test" in sys.argv:
        template = bundled_template()
        if not template:
            raise SystemExit(2)
        from openpyxl import load_workbook

        workbook = load_workbook(template, read_only=True, data_only=False)
        if not {"YN", "IC"}.issubset(workbook.sheetnames):
            raise SystemExit(3)
        workbook.close()
        raise SystemExit(0)

    if "--keygen" in sys.argv or "--admin" in sys.argv:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        dlg = AdminLicenseDialog(root)
        dlg.wait_window()
        root.destroy()
        sys.exit(0)

    # Check hardware license activation before running
    if not ensure_authorized():
        sys.exit(0)

    C12App().mainloop()


def main_free() -> None:
    """Phiên bản không kiểm tra bản quyền, mở lên là xài luôn ngay lập tức."""
    if "--self-test" in sys.argv:
        template = bundled_template()
        if not template:
            raise SystemExit(2)
        from openpyxl import load_workbook

        workbook = load_workbook(template, read_only=True, data_only=False)
        if not {"YN", "IC"}.issubset(workbook.sheetnames):
            raise SystemExit(3)
        workbook.close()
        raise SystemExit(0)

    C12App(unrestricted=True).mainloop()


if __name__ == "__main__":
    main()
