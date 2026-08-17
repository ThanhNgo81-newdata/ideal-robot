# MEP Translator (bản Python)

Ứng dụng dịch tài liệu kỹ thuật MEP (PDF, Word, Excel, Ảnh) chạy trên máy bạn
bằng Python + Streamlit, gọi trực tiếp Anthropic API từ phía server nên
**không còn lỗi "Failed to fetch"** như bản HTML/artifact chạy trong trình duyệt.

## 1. Vì sao bản trước bị lỗi "Failed to fetch"?

Bản HTML trước gọi thẳng `api.anthropic.com` bằng JavaScript trong trình
duyệt. Cách gọi đó chỉ được phép khi chạy **bên trong artifact của claude.ai**
(Anthropic mở proxy riêng cho môi trường đó). Nếu tải file .html về máy và mở
trực tiếp (hoặc mở ở nơi khác), trình duyệt sẽ chặn request do chính sách CORS
→ báo lỗi "Failed to fetch". Bản Python này gọi API từ phía server (không qua
trình duyệt) nên không gặp giới hạn đó, và chạy độc lập không cần claude.ai.

## 2. Cài đặt

Yêu cầu: Python 3.9 trở lên.

```bash
# Tạo môi trường ảo (khuyến khích)
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# Cài thư viện
pip install -r requirements.txt
```

## 3. Lấy Anthropic API key

1. Vào **console.anthropic.com** (hoặc platform.claude.com), đăng ký tài khoản.
2. Vào **Settings → Billing**, thêm phương thức thanh toán (API tính phí theo
   token, không có gói miễn phí vĩnh viễn, nhưng tài khoản mới thường có
   credit dùng thử).
3. Vào **API Keys → Create Key**, đặt tên, tạo key.
4. Key hiện dạng `sk-ant-...` — copy ngay vì chỉ hiện **một lần**.

## 4. Chạy ứng dụng

```bash
streamlit run mep_translator.py
```

Trình duyệt sẽ tự mở tại `http://localhost:8501`. Dán API key vào ô ở
thanh bên trái, chọn ngôn ngữ nguồn/đích, tải tệp lên và bấm **Dịch tài liệu**.

## 5. Tính năng

- **Định dạng hỗ trợ:** PDF, Word (.docx), Excel (.xlsx/.xls), ảnh (JPG/PNG/WEBP).
- **Cặp ngôn ngữ:** Anh↔Việt, Trung→Anh, Trung→Việt và các chiều khác
  (Anh/Việt/Trung, có chế độ tự động nhận diện ngôn ngữ nguồn).
- **Chuyên ngành MEP:** giữ nguyên mã thiết bị (AHU-01, FCU...), tiêu chuẩn
  (TCVN, QCVN, ASHRAE, NFPA...), đơn vị kỹ thuật (kW, CFM, Pa...); dùng thuật
  ngữ MEP chuẩn ngành Cơ - Điện - Nước.
- **Excel:** dịch từng ô, giữ nguyên cấu trúc bảng, tải về file .xlsx đã dịch —
  phù hợp để dịch bảng thống kê thiết bị (schedule).
- **PDF/Word/Ảnh:** trích xuất văn bản, dịch theo đoạn, xem song song bản
  gốc/bản dịch, tải về .txt.

## 6. Lưu ý

- PDF dạng scan ảnh (không có lớp văn bản) sẽ không trích xuất được bằng
  pypdf — trong trường hợp đó, chụp ảnh trang cần dịch và tải lên như ảnh.
- API key được lưu tạm trong phiên làm việc của trình duyệt (session), không
  ghi ra file — mỗi lần mở lại app cần nhập lại, hoặc bạn có thể sửa code để
  đọc từ biến môi trường `ANTHROPIC_API_KEY` nếu muốn tiện hơn.
- Tài liệu rất dài sẽ được chia nhỏ để dịch, có thể mất từ vài chục giây đến
  vài phút tùy độ dài.
