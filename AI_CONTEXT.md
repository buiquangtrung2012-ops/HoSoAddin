# BỐI CẢNH DỰ ÁN CHO AI (AI CONTEXT)

> **LƯU Ý DÀNH CHO AI (LLMs):** Khi bắt đầu một phiên làm việc mới trên dự án này, hãy đọc kỹ toàn bộ nội dung file này để nắm bắt cấu trúc, logic cốt lõi, và lịch sử giải quyết vấn đề trước khi đề xuất thay đổi mã nguồn.

## 1. TỔNG QUAN DỰ ÁN
Đây là một Office Add-in (dành cho Microsoft Word) phục vụ việc tự động hóa quá trình điền thông tin, tạo bảng biểu, nén/tách tệp hồ sơ dự án thi công.
- **UI Framework**: Vanilla JS, HTML5, Tailwind CSS (giao diện Solid, đổ bóng nhẹ, không dùng Glassmorphism), Lucide Icons.
- **API tương tác**: Office.js (Word API).
- **Thư viện phụ trợ**: JSZip (dùng để nén/tách file gốc Document.xml), File System Access API (để chọn và ghi file vào thư mục người dùng máy tính).

## 2. CẤU TRÚC TỆP QUAN TRỌNG
- **`taskpane.html` / `taskpane.js`**: File giao diện và file logic chính. Quản lý toàn bộ State (dự án, nhân sự, máy móc, vật liệu, thí nghiệm). Xử lý các sự kiện click nút (Cập nhật, Tách, Nhập dữ liệu).
- **`word_service.js`**: Class chứa các hàm thao tác trực tiếp với file Word. Đọc/ghi Content Controls, tìm Bookmark, xuất bảng tự động, xử lý giải nén/nén file DOCX dưới dạng nhị phân để tách file.
- **`storage_service.js`**: Dùng `Office.context.document.settings` để lưu cấu hình JSON vào file Word, và dùng `IndexedDB` để lưu trữ các `FileSystemDirectoryHandle` (thư mục xuất file).

## 3. LOGIC CỐT LÕI CẦN TUÂN THỦ
1. **Đồng bộ Dữ liệu (Sync Data)**:
   - Dữ liệu text đơn giản (Tên dự án, Gói thầu...) được đồng bộ vào các **Content Controls** dựa trên `Tag` hoặc `Alias`.
   - Dữ liệu dạng danh sách (Nhân sự, Máy móc...) được chèn tự động dưới dạng bảng tại vị trí của các **Bookmarks** (ví dụ: `bmNhanSu`, `bmMayMoc`).
2. **File System Access API (Ghi file vào máy)**:
   - Mọi thao tác lưu file qua `showDirectoryPicker` đều yêu cầu quyền truy cập. 
   - Add-in đã được cài đặt hàm `verifyPermission()` để **bắt buộc kiểm tra lại quyền** mỗi khi lấy Directory Handle từ IndexedDB (sau khi reload máy). Nếu mất quyền, hệ thống sẽ hiện hộp thoại xin cấp lại hoặc bắt người dùng chọn lại thư mục.
3. **Chế độ xuất file (Export Mode)**:
   - `master`: Cập nhật toàn bộ thông tin và xuất ra 1 file Word duy nhất. Có tùy chọn `useProjectNameFolder` để đặt tên file xuất ra là `"1. Ho so dau vao.docx"` nằm trong thư mục có Tên Dự Án.
   - `split`: Tách từng phân đoạn hồ sơ thành các file `.docx` riêng lẻ dựa trên các Bookmark đánh dấu trước, sau đó nén lại thành file `.zip` (bằng JSZip) hoặc thả trực tiếp vào thư mục máy tính.
4. **Xử lý chuỗi địa danh hành chính**:
   - Hàm `WordService.xuLyLayTenDuAn` được thiết kế để trích xuất địa danh (Xã, Huyện, Tỉnh). **Đặc biệt lưu ý**: Hàm đã được tối ưu để giữ lại khoảng trắng giữa các từ (VD: `Van Hoa_Ba Vi` thay vì `VanHoa_BaVi`).
5. **Bảng chữ ký Liên Danh**:
   - Có cơ chế tạo bảng chữ ký động tùy theo số lượng nhà thầu liên danh (từ 2 cột, 3 cột...). 
   - Thuật toán có heuristic để không ghi đè nhầm bảng ký tên lên Header/Footer của văn bản.

## 4. NHỮNG LỖI HÓC BÚA ĐÃ GIẢI QUYẾT (BUGS KILLED)
- 🐛 **Lỗi treo "Đang xuất bộ hồ sơ tổng..." (ở 30-50%)**: Do `FileSystemDirectoryHandle` lưu trong IndexedDB bị mất quyền Write sau khi reset trang. Đã fix bằng `verifyPermission()`. 
- 🐛 **Lỗi ghi đè Header khi chèn chữ ký**: Do logic tìm bảng mặc định bị nhầm sang bảng của Header. Đã fix bằng cách định vị thông qua bookmark `bmKyLienDanh`.
- 🐛 **Giao diện bảng bị méo/thiếu cân đối**: Đã fix bằng cách thiết lập tỷ lệ `%` trực tiếp trong mã (`wMap`) ở `taskpane.js` cho các cột: Tên thiết bị hẹp lại, Chủ sở hữu rộng ra, Chuyên ngành rộng ra để chứa các tên dài không bị vỡ dòng.

## 5. QUY ƯỚC LÀM VIỆC DÀNH CHO AI
1. Không sử dụng TailwindCSS version cũ, luôn tuân thủ class Tailwind chuẩn.
2. Không tùy ý xóa các comments giải thích logic hiện tại.
3. Cập nhật mã phiên bản (version string dạng `vDDMMYYYY.HHMM`) tại `taskpane.html` và `taskpane.js` mỗi khi có thay đổi.
4. Sau khi thay đổi code thành công, nếu nhận được lệnh thì sử dụng file `GitHub_Automation.ps1` để tự động đẩy code lên kho.

---
*Cập nhật lần cuối: Xem lịch sử Git commit.*
