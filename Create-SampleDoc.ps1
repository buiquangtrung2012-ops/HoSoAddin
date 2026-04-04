$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Add()

$selection = $word.Selection

$selection.TypeText("DỰ ÁN: ")
$range1 = $selection.Range
$doc.Bookmarks.Add("bmTenDuAn", $range1)
$selection.TypeParagraph()

$selection.TypeText("1. DANH SÁCH NHÂN SỰ")
$selection.TypeParagraph()
$table1 = $doc.Tables.Add($selection.Range, 2, 5)
$table1.Borders.Enable = $true
$table1.Cell(1, 1).Range.Text = "STT"
$table1.Cell(1, 2).Range.Text = "Họ và Tên"
$table1.Cell(1, 3).Range.Text = "Chức vụ"
$table1.Cell(1, 4).Range.Text = "Chuyên ngành"
$table1.Cell(1, 5).Range.Text = "Ghi chú"
$doc.Bookmarks.Add("bmNhanSu", $table1.Range)

$selection.Start = $doc.Content.End
$selection.TypeParagraph()
$selection.TypeText("2. DANH SÁCH THIẾT BỊ")
$selection.TypeParagraph()
$table2 = $doc.Tables.Add($selection.Range, 2, 4)
$table2.Borders.Enable = $true
$table2.Cell(1, 1).Range.Text = "STT"
$table2.Cell(1, 2).Range.Text = "Tên Thiết Bị"
$table2.Cell(1, 3).Range.Text = "Số hiệu"
$table2.Cell(1, 4).Range.Text = "Tình trạng"
$doc.Bookmarks.Add("bmMayMoc", $table2.Range)

$path = "C:\Users\buiqu\.gemini\antigravity\scratch\WordHoSoAddin\File_Test_Addin.docx"
$doc.SaveAs([ref]$path)
Write-Host "File test đã được tạo tại: $path"
# $word.Quit()
