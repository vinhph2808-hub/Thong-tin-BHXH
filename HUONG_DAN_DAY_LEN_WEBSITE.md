# HƯỚNG DẪN ĐẨY ỨNG DỤNG LÊN WEBSITE (DÙNG TRỰC TIẾP QUA LINK)

Hệ thống đã được thiết kế sẵn bản **Web App** (chạy bằng Streamlit). Người dùng chỉ cần vào đường link web bằng trình duyệt (Chrome, Cốc Cốc, Safari, Edge...), kéo thả file PDF/ảnh vào là xem được số liệu và bấm nút tải ngay file Excel.

Có 2 giải pháp tối ưu nhất để đưa lên website:

---

## 🌟 GIẢI PHÁP 1: Đưa lên Streamlit Community Cloud (Khuyên dùng - Chạy 24/7)

### Ưu điểm vượt trội:
- **Chạy 24/7 độc lập:** Máy tính của bạn tắt thì người khác vẫn vào website sử dụng bình thường.
- **Tự động cập nhật (CI/CD):** Mỗi khi bạn cập nhật code trên máy tính và bấm đẩy lên GitHub, máy chủ Cloud sẽ **tự động cập nhật phiên bản mới sau 10 - 20 giây**.
- **Miễn phí 100%:** Có đường link cố định rất đẹp dạng: `https://ten-ban-chon.streamlit.app`.

### Các bước thực hiện (chỉ làm 1 lần ban đầu):
1. **Tạo kho chứa (Repository) trên GitHub:**
   - Truy cập [github.com](https://github.com) và tạo một tài khoản (nếu chưa có).
   - Bấm nút **New repository** (đặt tên ví dụ: `xuat-thong-bao-bhxh-c12`), chọn chế độ **Public** hoặc **Private**.
2. **Đẩy mã nguồn lên GitHub:**
   - Các file cần thiết đã được chuẩn bị đầy đủ sẵn sàng gồm: `web_app.py`, `requirements.txt`, `packages.txt`, thư mục `src/`, và file mẫu `THÔNG BÁO BXH.xlsx`.
3. **Kết nối lên Streamlit Cloud:**
   - Truy cập [share.streamlit.io](https://share.streamlit.io) và đăng nhập bằng tài khoản GitHub vừa tạo.
   - Bấm **Create app** -> Chọn Repository của bạn.
   - Tại mục **Main file path**, chọn: `web_app.py`.
   - Bấm **Deploy!**
   - Sau khoảng 1-2 phút, bạn sẽ nhận được đường link chính thức (ví dụ: `https://c12-bhxh.streamlit.app`) để gửi cho mọi người sử dụng.

---

## 🚀 GIẢI PHÁP 2: Phát Web trực tiếp từ máy của bạn (Cloudflare Tunnel)

### Ưu điểm:
- Không cần tạo tài khoản GitHub hay cấu hình Cloud phức tạp.
- Cập nhật tức thì: Bạn sửa code trên máy là người dùng bên ngoài thấy thay đổi ngay lập tức.
- Hoàn toàn miễn phí và bảo mật tuyệt đối qua Cloudflare.

### Cách chạy:
1. Mở file [Chay_Web_App.bat](file:///d:/Vinh.ph/AI/Codex/Thông%20báo%20BHXH/Chay_Web_App.bat) để khởi chạy máy chủ Web nội bộ.
2. Dùng công cụ Cloudflare Tunnel để tạo đường link công khai dạng `https://...trycloudflare.com` gửi cho mọi người.
*(Lưu ý: Cách này chỉ truy cập được khi máy tính của bạn đang mở).*
