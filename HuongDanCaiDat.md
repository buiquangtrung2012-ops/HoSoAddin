# Hướng dẫn Cài đặt Add-in "Quản Lý Hồ Sơ" trong Word

Chào bạn! Đây là hướng dẫn từng bước để bạn có thể chạy thử nghiệm Add-in vừa được chuyển đổi từ VBA sang Office JS (JavaScript) ngay trong phần mềm Word của mình.

---

### Bước 1: Khởi động Server cục bộ bằng VS Code
Vì Office Add-in là một ứng dụng web chạy bên trong Word, bạn cần "mở" nó trên một trình duyệt (server) trước khi Word có thể đọc được.

1.  Mở thư mục `WordHoSoAddin` bằng **Visual Studio Code**.
2.  Nếu bạn đã cài extension **Live Server**, hãy nhấn vào nút **"Go Live"** ở góc dưới cùng bên phải của VS Code.
3.  Lưu ý: Live Server thường chạy ở cổng **5500**. Hãy đảm bảo địa chỉ là `http://localhost:5500/taskpane.html` (Trùng với thiết lập trong file `manifest.xml`).

---

### Bước 2: Thiết lập Thư mục tin cậy (Trusted Catalog) trong Word
Đây là cách Word nhận diện Add-in của bạn mà không cần đưa lên Store.

1.  **Chia sẻ thư mục**: 
    - Nhấn chuột phải vào thư mục `WordHoSoAddin` -> chọn **Properties**.
    - Sang thẻ **Sharing** -> Chọn **Share...** -> Chọn chính mình và nhấn **Share**.
    - Copy lại đường dẫn mạng (ví dụ: `\\Tên-Máy-Tính\WordHoSoAddin`).
2.  **Cấu hình Word**:
    - Mở **Microsoft Word** -> Vào menu **File** -> **Options**.
    - Chọn **Trust Center** -> **Trust Center Settings...**
    - Chọn **Trusted Add-in Catalogs**.
    - Tại ô **Catalog Url**, dán đường dẫn mạng bạn vừa copy -> Nhấn **Add catalog**.
    - Tích vào ô **Show in Menu**.
    - Nhấn **OK** và khởi động lại Word.

---

### Bước 3: Nạp (Sideload) Add-in vào văn bản
1.  Trong Word, vào thẻ **Insert** (hoặc **Home** tùy phiên bản) -> Chọn **Add-ins**.
2.  Chọn tab **SHARED FOLDER** ở phía trên.
3.  Bạn sẽ thấy Add-in **Quản Lý Hồ Sơ Dự Án** hiện ra. Nhấn **Add**.
4.  Bảng điều khiển (Task Pane) hiện đại sẽ xuất hiện bên phải màn hình Word.

---

### Ghi chú quan trọng
- **Https**: Nếu Word yêu cầu Https, bạn có thể cần thiết lập chứng chỉ SSL cho Live Server hoặc cấu hình Word để chấp nhận tệp `http://localhost`. Tuy nhiên, thông thường Word trên Desktop cho phép chạy `localhost` ở dạng `http` trong giai đoạn phát triển.
- **Dữ liệu**: Hãy thử nhấn nút "Đồng bộ dữ liệu" để xem Add-in giao tiếp với Word như thế nào!

### 6. QUAN TRỌNG: Sử dụng hàng ngày (Macro-Free)
Giờ đây bạn đã chuyển sang hệ thống Add-in hiện đại, bạn **KHÔNG CẦN** tệp `.docm` (chứa Macro) nữa:
1.  Mở tệp `.docm` của bạn, chọn **File > Save As**.
2.  Tại phần **File Type**, hãy chọn **Word Document (*.docx)**.
3.  Lưu tệp mới này và xóa tệp `.docm` cũ đi.
4.  Từ nay, bạn chỉ cần mở tệp `.docx` sạch sẽ này, Add-in sẽ luôn tự động hiện lên bên phải để phục vụ bạn.

**Chúc mừng bạn đã nâng cấp thành công hệ thống "Hồ sơ dự án" lên tầm cao mới!**
Chúc bạn thành công! Nếu gặp khó khăn ở bước nào, hãy nhắn cho tôi biết ngay.
