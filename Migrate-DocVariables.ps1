# SCRIPT CHUYỂN ĐỔI DOCVARIABLE SANG CONTENT CONTROLS (Tự động)
# Yêu cầu: Đang mở tệp Word cần chuyển đổi

write-host "Đang kết nối tới Word đang mở..." -ForegroundColor Cyan

try {
    $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
} catch {
    write-host "KHÔNG TÌM THẤY WORD! Vui lòng mở tệp tài liệu Word chứa mẫu lên trước." -ForegroundColor Red
    exit
}

if ($word.Documents.Count -eq 0) {
    write-host "Không có tài liệu nào đang mở. Vui lòng mở tệp Word mẫu của bạn." -ForegroundColor Red
    exit
}

$doc = $word.ActiveDocument
$fields = $doc.Fields
$count = $fields.Count
$converted = 0

write-host "Tìm thấy $count trường (Fields) trong văn bản. Tiến hành kiểm tra và chuyển đổi..." -ForegroundColor Yellow

# Lặp ngược vì chúng ta sẽ xóa field
for ($i = $count; $i -ge 1; $i--) {
    $field = $fields.Item($i)
    # wdFieldDocVariable = 64
    if ($field.Type -eq 64) {
        $code = $field.Code.Text.Trim()
        
        # Regex trích xuất tên biến (VD: DOCVARIABLE DuAn -> DuAn)
        if ($code -match "DOCVARIABLE\s+`"?([A-Za-z0-9_-]+)`"?") {
            $varName = $matches[1]
            $range = $field.Result
            $text = $range.Text
            
            # Xóa Field cũ
            $field.Delete()
            
            # Thay thế bằng Content Control (Rich Text = 1)
            $cc = $doc.ContentControls.Add(1, $range)
            $cc.Tag = $varName
            $cc.Title = $varName
            
            # Giữ nguyên chữ hoặc đặt giá trị trống
            if ($null -eq $text -or $text.Trim() -eq "") {
                $cc.Range.Text = " "
            } else {
                $cc.Range.Text = $text
            }
            
            $converted++
            write-host "Đã chuyển đổi thành công: Biến [$varName]" -ForegroundColor Green
        }
    }
}

$doc.Save()
write-host "==================================================" -ForegroundColor Cyan
write-host "HOÀN TẤT! Đã chuyển đổi $converted biến DocVariable sang Content Controls hiện đại." -ForegroundColor Cyan
write-host "==================================================" -ForegroundColor Cyan

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
