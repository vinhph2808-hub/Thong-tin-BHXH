@echo off
chcp 65001 >nul
title Hệ Thống Xuất Thông Báo BHXH C12 (Bản Web)
echo ======================================================================
echo    HỆ THỐNG XỬ LÝ & XUẤT THÔNG BÁO BHXH C12 - GIAO DIỆN WEB
echo ======================================================================
echo Đang khởi chạy máy chủ Web... Trình duyệt sẽ tự động mở!
echo Bạn cũng có thể mở thủ công tại địa chỉ: http://localhost:8501
echo ======================================================================
set PYTHONPATH=src
".venv\Scripts\streamlit.exe" run web_app.py --server.port 8501
pause
