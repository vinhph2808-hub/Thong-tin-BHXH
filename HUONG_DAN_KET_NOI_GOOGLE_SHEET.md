# 📊 HƯỚNG DẪN KẾT NỐI THEO DÕI MÁY TÍNH VÀO GOOGLE SHEET

Hệ thống đã được tích hợp sẵn tính năng tự động ghi nhận nhật ký (truy cập, xử lý file, tải kết quả) của mọi máy tính/thiết bị vào Google Sheet của bạn.

---

### Bước 1: Tạo bảng tính Google Sheet
1. Mở [Google Sheets](https://sheets.new) và tạo một trang tính mới.
2. Đặt tiêu đề cho các cột ở **Dòng 1**:
   * **Cột A**: `Thời gian`
   * **Cột B**: `Loại ứng dụng`
   * **Cột C**: `Tên thiết bị / Trình duyệt`
   * **Cột D**: `Người dùng / Địa chỉ IP`
   * **Cột E**: `Hành động thực hiện`
   * **Cột F**: `Chi tiết / Ghi chú`

---

### Bước 2: Dán mã Google Apps Script
1. Trên thanh menu của Google Sheet, bấm vào **Tiện ích mở rộng** (Extensions) > chọn **Apps Script**.
2. Xóa hết mã mặc định và dán đoạn mã sau vào:

```javascript
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    sheet.appendRow([
      data.timestamp || Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss"),
      data.app_type || "Không xác định",
      data.device || "Máy khách",
      data.user || "Ẩn danh",
      data.action || "Truy cập",
      data.details || ""
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

3. Bấm biểu tượng **Lưu** (💾 Save).

---

### Bước 3: Triển khai ứng dụng Web (Deploy)
1. Ở góc trên bên phải trang Apps Script, bấm nút **Triển khai** (Deploy) > chọn **Tùy chọn triển khai mới** (New deployment).
2. Bấm vào biểu tượng bánh răng ⚙️ bên cạnh "Chọn loại", chọn **Ứng dụng web** (Web app).
3. Cài đặt các thông số:
   * **Mô tả**: `Ghi log he thong C12`
   * **Thực thi dưới dạng** (Execute as): **Tôi** (Me)
   * **Ai có quyền truy cập** (Who has access): **Bất kỳ ai** (Anyone)  *(Bắt buộc chọn "Bất kỳ ai" để hệ thống web và app gửi dữ liệu vào mà không cần đăng nhập tài khoản Google)*.
4. Bấm **Triển khai** (Deploy) > Cấp quyền truy cập nếu Google yêu cầu.
5. Sao chép đường link **URL ứng dụng web** (dạng `https://script.google.com/macros/s/.../exec`).

---

### Bước 4: Nhập link vào hệ thống
Bạn có thể kết nối bằng 1 trong các cách sau:
* **Cách 1 (Trực tiếp trên Web)**: Mở trang web hệ thống, ở thanh bên trái bấm vào mục **📊 Cấu hình Google Sheet theo dõi**, dán link URL vừa copy vào và bấm **Lưu cấu hình**.
* **Cách 2 (File cấu hình dùng chung)**: Dán link URL vào file `google_sheet_config.json` trong thư mục gốc dự án:
```json
{
  "webhook_url": "https://script.google.com/macros/s/.../exec"
}
```
* **Cách 3 (Streamlit Cloud Secrets)**: Trên Streamlit Cloud > Settings > Secrets, thêm dòng:
```toml
GOOGLE_SHEET_WEBHOOK_URL = "https://script.google.com/macros/s/.../exec"
```

---
*Từ bây giờ, bất cứ khi nào có ai mở trang web hoặc ứng dụng desktop, xử lý thông báo C12 hay tải file, Google Sheet sẽ tự động nhảy thêm 1 dòng ghi nhận đầy đủ thời gian, thiết bị, IP và nội dung thao tác!*
