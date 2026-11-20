@echo off
chcp 65001 >nul
title CẤP MÃ BẢN QUYỀN - PHÊ DUYỆT THIẾT BỊ (ADMIN)

if exist "Xuat_Thong_Bao_BHXH_C12.exe" (
    start "" "Xuat_Thong_Bao_BHXH_C12.exe" --keygen
    exit /b 0
)

if exist "dist\Xuat_Thong_Bao_BHXH_C12.exe" (
    start "" "dist\Xuat_Thong_Bao_BHXH_C12.exe" --keygen
    exit /b 0
)

.venv\Scripts\python.exe -c "import tkinter as tk; from src.c12_tool.license_dialog import AdminLicenseDialog; r=tk.Tk(); r.withdraw(); d=AdminLicenseDialog(r); d.wait_window(); r.destroy()"
