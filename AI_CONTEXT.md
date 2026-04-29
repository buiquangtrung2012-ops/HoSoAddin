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
   - **Xử lý Bookmark bao trùm bảng**: Logic `xuatBang` sử dụng `compareLocationWith` với tham số `"Equal"` để nhận diện chính xác các Bookmark ôm trọn toàn bộ bảng, tránh việc bỏ sót bảng khi cập nhật.
2. **File System Access API (Ghi file vào máy)**:
   - Mọi thao tác lưu file qua `showDirectoryPicker` đều yêu cầu quyền truy cập. 
   - Add-in đã được cài đặt hàm `verifyPermission()` để **bắt buộc kiểm tra lại quyền** mỗi khi lấy Directory Handle từ IndexedDB (sau khi reload máy). Nếu mất quyền, hệ thống sẽ hiện hộp thoại xin cấp lại hoặc bắt người dùng chọn lại thư mục.
3. **Chế độ xuất file (Export Mode)**:
   - `master`: Cập nhật toàn bộ thông tin và xuất ra 1 file Word duy nhất. Có tùy chọn `useProjectNameFolder` để đặt tên file xuất ra là `"1. Ho so dau vao.docx"` nằm trong thư mục có Tên Dự Án.
   - `split`: Tách từng phân đoạn hồ sơ thành các file `.docx` riêng lẻ dựa trên các Bookmark đánh dấu trước (`TT_`, `SPLIT_`), sau đó nén lại thành file `.zip` (bằng JSZip) hoặc thả trực tiếp vào thư mục máy tính.
4. **Xử lý chuỗi địa danh hành chính**:
   - Hàm `WordService.xuLyLayTenDuAn` được thiết kế để trích xuất địa danh (Xã, Huyện, Tỉnh). **Đặc biệt lưu ý**: Hàm đã được tối ưu để giữ lại khoảng trắng giữa các từ (VD: `Van Hoa_Ba Vi` thay vì `VanHoa_BaVi`).
5. **Bảng mẫu và Chữ ký**:
   - Bảng ký tên (`bmKyLienDanh`) mặc định được chèn với định dạng **No Border** (không viền) và căn lề chuẩn (Nơi nhận in nghiêng bên trái, Đơn vị ký in đậm ở giữa).
   - Sử dụng phương thức `insertTable("After")` và truy cập ô trực tiếp qua `table.getCell(row, col)` để tránh lỗi `InvalidArgument` và tăng tính ổn định trên các phiên bản Word khác nhau.

## 4. NHỮNG LỖI HÓC BÚA ĐÃ GIẢI QUYẾT (BUGS KILLED)
- 🐛 **Lỗi không điền dữ liệu bảng Vật liệu**: Do không khớp từ khóa tìm kiếm dự phòng ("Tên vật tư" vs "Tên vật liệu") và Bookmark bao trùm bảng không được nhận diện. Đã fix bằng cách bổ sung từ khóa và toán tử so sánh `"Equal"`.
- 🐛 **Lỗi InvalidArgument khi chèn bảng**: Do việc gán Style và truy cập Rows theo Index không ổn định. Đã fix bằng cách dùng `After` insertion và `getCell`.
- 🐛 **Lỗi treo "Đang xuất bộ hồ sơ tổng..."**: Do `FileSystemDirectoryHandle` mất quyền Write. Đã fix bằng `verifyPermission()`. 
- 🐛 **Giao diện bảng bị méo**: Đã fix bằng cách thiết lập tỷ lệ `%` thủ công cho từng loại bảng trong `taskpane.js`.

## 5. CẢI TIẾN GIAO DIỆN (UI/UX)
- **Template Creator**: Sử dụng lưới 3 cột (grid-cols-3) cho các nút chèn mẫu. Các tiêu đề phân đoạn (Content Controls, Bookmarks, Split Markers) được làm nổi bật bằng dải màu nền Indigo và thanh lề trái đậm.
- **Tab Dự án**: Loại bỏ các icon bên trái để tăng tối đa độ rộng cho các trường nhập liệu (textarea), giúp hiển thị nội dung dài tốt hơn.
- **Tính đồng nhất**: Sửa lỗi hiển thị tiêu đề "XUATBAN" thành "CÀI ĐẶT".

## 6. QUY ƯỚC LÀM VIỆC DÀNH CHO AI
1. Luôn ưu tiên dùng Tailwind CSS Solid (không Glassmorphism).
2. Không tùy ý xóa các comments giải thích logic hiện tại.
3. Cập nhật mã phiên bản (version string dạng `vDDMMYYYY.HHMM`) tại `taskpane.html` và `taskpane.js` mỗi khi có thay đổi.
4. Sau khi thay đổi code thành công, sử dụng file `GitHub_Automation.ps1` để đẩy code lên kho.

---
*Cập nhật lần cuối: 29/04/2026*
