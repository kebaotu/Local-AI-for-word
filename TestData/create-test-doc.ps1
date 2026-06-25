# Script tạo tài liệu test sample-errors.docx
# Chạy trên máy có Word cài đặt

$word = New-Object -ComObject Word.Application
$word.Visible = $false

$doc = $word.Documents.Add()
$sel = $word.Selection

# Nội dung có lỗi cố ý
$content = @"
Công ty  chúng tôi đã đã tiến hành kiểm tra văn bằng này .

Văn bản được soạn ngày 20/05/2025. Đề nghị trả lời trước ngày 15/05/2025.

Chính sách nhập khẩu của Nhật Bản được phân tích trong mục này. Tuy nhiên, chính sách thuế của Hàn Quốc nêu trên cần được đánh giá thêm.

1. Nội dung thứ nhất
2. Nội dung thứ hai
4. Nội dung thứ tư

Theo mục 7 của văn bản này, các bên phải thực hiện đầy đủ nghĩa vụ. Tuy nhiên văn bản không có mục 7.

Tổng kinh phí là 1.000.000.000 đồng. Trong phần sau, kinh phí thực hiện được nêu là 100.000.000 đồng.

Bộ Tài chính đã ban hành văn bản. Sau đó, Bộ tài Chính tiếp tục hướng dẫn thực hiện.

Hợp đồng này có hiệu lực vĩnh viễn và Bên A không chịu bất kỳ trách nhiệm nào trong mọi trường hợp.

[Tên công ty] có trách nhiệm gửi báo cáo trước ngày {{Ngày báo cáo}}.
"@

$sel.TypeText($content)

# Thêm comments
$doc.Comments.Add($doc.Range(0, 10), "Cần làm rõ thời hạn phản hồi.")
$doc.Comments.Add($doc.Range(50, 100), "Điều khoản trách nhiệm đang quá bất lợi.")
$doc.Comments.Add($doc.Range(200, 250), "Kiểm tra lại số tiền tổng cộng.")

# Lưu file
$savePath = "$PSScriptRoot\sample-errors.docx"
$doc.SaveAs2($savePath)
$doc.Close()
$word.Quit()

Write-Host "Đã tạo: $savePath"
