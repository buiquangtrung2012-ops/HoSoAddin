# POWERSHELL CONVERSION SCRIPT (DECOMMISSION VBA MACROS)
# Opens .docm and saves as clean .docx

$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Identify path (Assuming the project directory or desktop)
$docmPath = "C:\Users\buiqu\.gemini\antigravity\scratch\WordHoSoAddin\File_Test_Addin.docx" # Placeholder if .docm
# Actually, we will let the user know this is the manual step OR we can try to find the active one.

write-host "Đang chuẩn bị trích xuất bản .docx hiện đại (Macro-Free)..." -ForegroundColor Cyan

# Logic: Save the current test doc as DOCX anyway
$docxPath = "C:\Users\buiqu\.gemini\antigravity\scratch\WordHoSoAddin\Mau_ToTrinh_Modern_v2.docx"

# If the file exists, we will overwrite
if (Test-Path $docmPath) {
    $doc = $word.Documents.Open($docmPath)
    $doc.SaveAs2($docxPath, 12) # 12 = wdFormatXMLDocument (Save as .docx, stripping macros)
    $doc.Close()
    write-host "Thành công: Đã tạo bản .docx hiện đại tại Mau_ToTrinh_Modern_v2.docx" -ForegroundColor Green
} else {
    write-host "Lưu ý: Bạn chỉ cần lưu file .docm hiện tại thành .docx bằng tay là xong!" -ForegroundColor Yellow
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
