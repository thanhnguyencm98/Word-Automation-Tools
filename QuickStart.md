# Quick Start Guide: Word Citation Automation Tools

Hướng dẫn sử dụng nhanh công cụ macro từ file `WordAutomationTools.docm`.

---

## 1. Chuẩn bị (Chỉ làm lần đầu)

1. Tải về hoặc mở file **`WordAutomationTools.docm`**.
2. Nếu Word hiển thị thanh màu vàng cảnh báo bảo mật: Bấm **Enable Content** (Bật nội dung).

---

## 2. Quy trình sử dụng hàng ngày

### Bước 1: Mở các tài liệu
1. Mở file công cụ **`WordAutomationTools.docm`**.
2. Mở file tài liệu Word (`.docx` hoặc `.docm`) mà bạn cần kiểm tra, rà soát citations.

### Bước 2: Chạy Macro
1. **Chuyển sang cửa sổ file `WordAutomationTools.docm`**:
   - Nhấn phím **Alt + F8** trên bàn phím.
   *(Lưu ý: Phải bấm Alt + F8 từ cửa sổ `WordAutomationTools.docm` thì Word mới hiển thị danh sách macro).*
2. Chọn một trong các macro:
   - **`RunAllChecks`**: Thực hiện toàn bộ 3 bước (gỡ link $\rightarrow$ bôi vàng Italic $\rightarrow$ bôi vàng Small Caps trong footnote). *(Khuyên dùng)*
   - **`RemoveAllHyperlinks`**: Chỉ gỡ bỏ hyperlinks.
   - **`HighlightItalicFootnotes`**: Chỉ bôi vàng chữ nghiêng trong footnote.
   - **`HighlightSmallCapsFootnotes`**: Chỉ bôi vàng chữ Small Caps trong footnote.
3. Bấm nút **Run**.

### Bước 3: Chọn tài liệu cần xử lý
- **Nếu bạn chỉ mở 1 file tài liệu:** Hộp thoại sẽ hỏi xác nhận:  
  `"Do you want to process the open target document: 'Ten_File.docx'?"` $\rightarrow$ Chọn **Yes**.
- **Nếu bạn đang mở nhiều file tài liệu cùng lúc:** Hộp thoại sẽ hiển thị danh sách đánh số:
  ```text
  Multiple target documents are open.
  Enter the number corresponding to the document you want to process:
  
  1. LegalBrief_v1.docx
  2. ResearchNote.docx
  ```
  Bạn chỉ cần nhập số (ví dụ: `1`) và nhấn **OK**.

### Bước 4: Xem kết quả
- Khi chạy xong, macro hiển thị hộp thoại thống kê chi tiết (số hyperlink đã gỡ, số đoạn text được highlight).
- Bấm **OK** và lưu lại file đã sửa.

---

## 3. Lưu ý an toàn
- Luôn lưu 1 bản sao lưu (backup copy) của tài liệu trước khi chạy gỡ hyperlink.
- File công cụ `WordAutomationTools.docm` được bảo vệ tự động, macro sẽ không bao giờ tự ý sửa đổi lên chính nó.