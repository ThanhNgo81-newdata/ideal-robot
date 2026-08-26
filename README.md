# MEP Translator Offline

Ứng dụng dịch tài liệu kỹ thuật MEP (Cơ – Điện – Nước) hỗ trợ **DOCX, PDF, Excel, TXT**.  
Chạy bằng mô hình Hugging Face (MarianMT) và có chế độ **song ngữ**. Ứng dụng cố gắng **giữ cấu trúc file gốc**: chèn bản dịch theo thứ tự phần tử hoặc thêm cột song song cho Excel/DOCX.

---

## Tính năng chính
- Dịch văn bản kỹ thuật giữa nhiều ngôn ngữ.
- Chế độ **song ngữ**: giữ bản gốc và hiển thị bản dịch.
- Cố gắng **giữ cấu trúc file gốc**:
  - DOCX: chèn bản dịch ngay sau đoạn gốc; tùy chọn bảng 2 cột cho layout song song.
  - Excel: chèn cột dịch kế bên và sao chép style cơ bản.
  - PDF: xuất ra DOCX theo thứ tự block; layout phức tạp có thể cần công cụ chuyên dụng.
  - TXT: giữ dòng gốc và chèn dòng dịch ngay sau.
- **Cache model/tokenizer** để giảm thời gian load và giảm tải CPU/GPU khi deploy.

---

## Yêu cầu hệ thống
- Python 3.9+ (khuyến nghị 3.10+)
- Thư viện được liệt kê trong `requirements.txt`.

**Lưu ý quan trọng**: các mô hình MarianMT yêu cầu `sentencepiece` và đôi khi `sacremoses`. Đảm bảo `requirements.txt` có hai gói này để tránh lỗi `ImportError` khi load tokenizer.

---

## Cài đặt và chạy local

1. Clone repo:
```bash
git clone https://github.com/<username>/<repo>.git
cd <repo>
