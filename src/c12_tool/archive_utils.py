from __future__ import annotations

import os
import shutil
import subprocess
import tarfile
import zipfile
from pathlib import Path

# Các định dạng tài liệu C12 hỗ trợ
SUPPORTED_DOC_EXTS = {
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".tiff",
    ".webp",
}

# Các định dạng file nén hỗ trợ
ARCHIVE_EXTS = {
    ".zip",
    ".rar",
    ".7z",
    ".tar",
    ".gz",
    ".tgz",
    ".bz2",
    ".tbz2",
    ".xz",
    ".txz",
}


def is_archive_file(path_or_name: str | Path) -> bool:
    """Kiểm tra xem tập tin có phải là file nén không."""
    p = Path(path_or_name)
    suffix = p.suffix.lower()
    return suffix in ARCHIVE_EXTS


def extract_archive(archive_path: str | Path, extract_dir: str | Path) -> list[Path]:
    """
    Giải nén tập tin nén (.zip, .rar, .7z, .tar, .gz...) vào thư mục chỉ định.
    Hỗ trợ tiếng Việt Unicode, xử lý đệ quy file nén lồng nhau,
    và trả về danh sách các đường dẫn file tài liệu C12 (.pdf, ảnh) tìm được.
    """
    archive_p = Path(archive_path)
    dest_p = Path(extract_dir)
    dest_p.mkdir(parents=True, exist_ok=True)
    suffix = archive_p.suffix.lower()

    extracted = False

    # 1. Thử giải nén bằng zipfile tiêu chuẩn
    if suffix == ".zip" or zipfile.is_zipfile(archive_p):
        try:
            with zipfile.ZipFile(archive_p, "r") as zf:
                for member in zf.infolist():
                    target_name = member.filename
                    # Xử lý encoding tiếng Việt nếu nén bằng winrar/windows cp437
                    if not (member.flag_bits & 0x800):
                        try:
                            target_name = member.filename.encode("cp437").decode("utf-8")
                        except Exception:
                            try:
                                target_name = member.filename.encode("cp437").decode("cp1258")
                            except Exception:
                                target_name = member.filename

                    target_file = dest_p / target_name
                    # Bảo vệ chống Zip Slip (vượt ra ngoài thư mục đích)
                    try:
                        if not target_file.resolve().is_relative_to(dest_p.resolve()):
                            continue
                    except Exception:
                        pass

                    if member.is_dir() or target_name.endswith("/"):
                        target_file.mkdir(parents=True, exist_ok=True)
                    else:
                        target_file.parent.mkdir(parents=True, exist_ok=True)
                        with zf.open(member) as src_f, open(target_file, "wb") as dst_f:
                            shutil.copyfileobj(src_f, dst_f)
            extracted = True
        except Exception:
            extracted = False

    # 2. Thử giải nén bằng tarfile tiêu chuẩn
    if not extracted and suffix in {".tar", ".gz", ".tgz", ".bz2", ".tbz2", ".xz", ".txz"}:
        try:
            with tarfile.open(archive_p, "r:*") as tf:
                # filter data nếu Python >= 3.12 để an toàn
                if hasattr(tarfile, "data_filter"):
                    tf.extractall(dest_p, filter="data")
                else:
                    tf.extractall(dest_p)
            extracted = True
        except Exception:
            extracted = False

    # 3. Thử giải nén bằng bsdtar hệ thống (có sẵn trên Windows 10/11 & Linux, đọc được zip, rar, 7z, tar...)
    if not extracted:
        tar_cmd = shutil.which("tar")
        if tar_cmd:
            try:
                res = subprocess.run(
                    [tar_cmd, "-xf", str(archive_p), "-C", str(dest_p)],
                    capture_output=True,
                    text=True,
                    timeout=90,
                )
                if res.returncode == 0:
                    extracted = True
            except Exception:
                pass

    # 4. Thử bằng 7-Zip CLI (7z hoặc 7za) nếu máy có cài đặt
    if not extracted:
        seven_zip = shutil.which("7z") or shutil.which("7za")
        if seven_zip:
            try:
                res = subprocess.run(
                    [seven_zip, "x", "-y", f"-o{dest_p}", str(archive_p)],
                    capture_output=True,
                    text=True,
                    timeout=90,
                )
                if res.returncode == 0:
                    extracted = True
            except Exception:
                pass

    # 5. Thử fallback shutil.unpack_archive
    if not extracted:
        try:
            shutil.unpack_archive(str(archive_p), str(dest_p))
            extracted = True
        except Exception:
            pass

    # Quét đệ quy tìm tất cả các file tài liệu hợp lệ
    found_docs: list[Path] = []
    if dest_p.exists():
        for item in sorted(dest_p.rglob("*")):
            if not item.is_file():
                continue
            # Bỏ qua file ẩn rác hệ thống (macOS ._ hoặc Windows Thumbs.db)
            if item.name.startswith("._") or item.name.startswith("Thumbs.db") or item.name.startswith(".DS_Store"):
                continue

            item_suffix = item.suffix.lower()
            if item_suffix in SUPPORTED_DOC_EXTS:
                found_docs.append(item)
            elif item_suffix in ARCHIVE_EXTS:
                # Đệ quy giải nén file nén con bên trong
                sub_unpacked = item.parent / f"_unpacked_{item.stem}"
                sub_docs = extract_archive(item, sub_unpacked)
                found_docs.extend(sub_docs)

    return found_docs
