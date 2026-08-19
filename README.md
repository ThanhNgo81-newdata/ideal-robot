# MEP Translator (bản Python) — v2

Ứng dụng dịch tài liệu kỹ thuật MEP (PDF lớn, Word, Excel, Ảnh/ảnh chụp màn
hình) chạy trên máy bạn bằng Python + Streamlit, gọi trực tiếp Anthropic API
từ phía server — không phụ thuộc claude.ai, không lỗi CORS/"Failed to fetch".

## 1. Cài đặt

Yêu cầu: Python 3.9 trở lên.

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Lấy Anthropic API key

1. Vào **console.anthropic.com** (hoặc platform.claude.com), đăng ký tài khoản.
2. Vào **Settings → Billing**, thêm phương thức thanh toán (API tính phí theo
   token, không có gói miễn phí vĩnh viễn, nhưng tài khoản mới thường có
   credit dùng thử).
3. Vào **API Keys → Create Key**, đặt tên, tạo key.
4. Key hiện dạng `sk-ant-...` — copy ngay vì chỉ hiện **một lần**.

## 3. Chạy ứng dụng

```bash
streamlit run mep_translator.py
```

Trình duyệt tự mở tại `http://localhost:8501`. Dán API key vào ô ở thanh bên
trái, chọn ngôn ngữ nguồn/đích, (tùy chọn) bật chế độ song ngữ, tải tệp lên
và bấm **Dịch tài liệu**.

## 4. Tính năng

- **Định dạng:** PDF (kể cả PDF nhiều trang), Word (.docx), Excel
  (.xlsx/.xls), ảnh và ảnh chụp màn hình (JPG/PNG/WEBP).
- **Đa ngôn ngữ:** Anh, Việt, Trung, Nhật, Hàn, Pháp, Đức, Thái, Tây Ban Nha,
  Nga, Indonesia, Mã Lai, Khmer, Lào, Hindi, Ả Rập, Bồ Đào Nha, Ý — dịch được
  theo **mọi chiều** giữa các ngôn ngữ này, có chế độ tự động nhận diện ngôn
  ngữ nguồn, và tùy chọn **"Khác (tự nhập)"** để gõ ngôn ngữ bất kỳ khác.
- **Chế độ song ngữ** (bật ở sidebar) — ví dụ cặp Anh-Nhật, Anh-Việt:
  - *Word:* chèn thêm đoạn dịch (in nghiêng) ngay sau mỗi đoạn gốc.
  - *Excel:* **chèn thêm 1 cột dịch ngay bên phải mỗi cột có chữ**, ngay
    trong sheet gốc — giống định dạng "Nội dung kiểm tra (VN) | Content (EN)"
    cạnh nhau mà kỹ sư MEP hay dùng trong checklist/specification song ngữ.
    Giữ nguyên toàn bộ merge, màu nền, độ rộng cột — **không** tạo sheet
    "đối chiếu" riêng. Dòng tiêu đề/banner merge rộng (spans nhiều cột) sẽ
    gộp gốc + dịch trong cùng ô (vì bản chất không tách cột được).
  - *PDF:* file Word xuất ra có cả đoạn gốc (in nghiêng) và đoạn dịch, theo
    đúng thứ tự trang.
  - *Ảnh:* văn bản gốc và bản dịch hiển thị xen kẽ trong kết quả.
- **Giữ nguyên cấu trúc file gốc:**
  - *Word:* dịch tại chỗ theo từng đoạn/ô bảng, giữ nguyên heading, bullet,
    numbering, bảng biểu, style gốc.
  - *Excel:* dịch từng ô, giữ nguyên sheet, định dạng, màu sắc, border,
    merged cells; công thức không bị thay đổi.
  - *PDF:* trích xuất đúng thứ tự đọc từng trang, dịch, rồi **tái dựng thành
    Word** giữ nguyên thứ tự trang/đoạn văn — cách ổn định nhất để tránh lỗi
    font khi hiển thị dấu tiếng Việt/chữ Hán/chữ Nhật.
  - *Ảnh/ảnh chụp màn hình:* Claude tự nhận diện chữ (OCR tích hợp), cố gắng
    giữ bố cục bảng bằng dấu `|`.
- **PDF lớn:** app hiện số trang, cho chọn khoảng trang cần dịch để kiểm soát
  thời gian và chi phí; cảnh báo nếu vượt quá 25 trang.
- **Chuyên ngành MEP:** giữ nguyên mã thiết bị (AHU-01, FCU...), tiêu chuẩn
  (TCVN, QCVN, ASHRAE, NFPA, ASME, SMACNA, IEC, JIS...), đơn vị kỹ thuật (kW,
  CFM, Pa...), dùng thuật ngữ chuẩn ngành Cơ - Điện - Nước.

## 5. Dùng chung với đồng nghiệp — API key vẫn ở phía bạn (tùy chọn)

Mặc định app bắt nhập API key thủ công (dành cho bạn tự dùng). Để đồng
nghiệp dùng chung mà **không thấy, không cần nhập key**:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Mở file vừa tạo, điền:

```toml
ANTHROPIC_API_KEY = "sk-ant-key-that-cua-ban"
APP_PASSWORD = "mat-khau-noi-bo"   # tùy chọn, nên đặt nếu chia sẻ ra ngoài LAN
```

File này đã có trong `.gitignore` — **không bao giờ commit lên GitHub**. Khi
đã cấu hình, sidebar tự ẩn ô nhập key, đồng nghiệp chỉ việc tải file và dịch.

**Chia sẻ trong mạng nội bộ (nhanh nhất):**
```bash
streamlit run mep_translator.py --server.address 0.0.0.0 --server.port 8501
```
Terminal in ra **Network URL** (`http://192.168.x.x:8501`) — gửi cho đồng
nghiệp cùng mạng LAN/wifi, máy bạn cần bật app suốt lúc họ dùng.

**Chia sẻ từ xa, chạy 24/7 (miễn phí):** push code (không kèm
`secrets.toml` thật) lên GitHub → vào **share.streamlit.io** → **New app** →
chọn repo và file `mep_translator.py` → mục **Advanced settings → Secrets**
dán nội dung `secrets.toml` của bạn vào đó → **Deploy**. Bạn nhận một URL
công khai, key vẫn chỉ nằm trong Secrets được mã hóa của Streamlit Cloud.

## 6. Lưu ý & giới hạn kỹ thuật

- PDF dạng scan ảnh (không có lớp văn bản) sẽ không trích xuất được — chụp
  ảnh trang đó và tải lên như ảnh (chế độ ảnh dùng OCR của Claude).
- PDF được xuất ra dưới dạng **Word tái dựng** (không sửa trực tiếp lên PDF
  gốc) để đảm bảo hiển thị đúng mọi ngôn ngữ trên mọi máy — layout theo đúng
  thứ tự đọc và ngắt trang, nhưng không giữ vị trí pixel-chính-xác như bản
  gốc (bảng phức tạp có thể cần chỉnh lại thủ công trong Word).
- Nếu không cấu hình `secrets.toml`, API key nhập tay chỉ lưu tạm trong
  phiên trình duyệt, không ghi ra file.
- PDF càng nhiều trang, số lượt gọi API càng nhiều → thời gian chờ và chi phí
  tăng theo. Với PDF rất lớn (hàng trăm trang), nên dùng tính năng chọn
  khoảng trang để dịch từng phần.
