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

- **Định dạng hỗ trợ:** PDF (kể cả PDF nhiều trang/PDF lớn), Word (.docx),
  Excel (.xlsx/.xls), ảnh (JPG/PNG/WEBP).
- **Đa ngôn ngữ như Google Translate:** Anh, Việt, Nhật, Trung (giản thể/phồn
  thể), Hàn, Thái, Pháp, Đức, Tây Ban Nha, Bồ Đào Nha, Ý, Nga, Indonesia, Mã
  Lai, Khmer, Lào, Hindi, Ả Rập... dịch được **theo mọi chiều** giữa các ngôn
  ngữ này (không chỉ Anh/Việt/Trung), có chế độ **tự động nhận diện ngôn ngữ
  nguồn**, và tùy chọn **"Khác... (tự nhập)"** để gõ tên một ngôn ngữ bất kỳ
  không có trong danh sách.
- **Chế độ song ngữ (bật/tắt ở sidebar):** kết quả hiển thị đối chiếu đoạn
  gốc và bản dịch xen kẽ (khuôn `[GỐC]` / `[DỊCH]`) thay vì chỉ hiện bản dịch
  — tiện đối chiếu thuật ngữ khi làm hồ sơ thầu, specification song ngữ.
  Với Excel, chế độ này tạo thêm 1 sheet **"Song ngữ (đối chiếu)"** liệt kê
  đầy đủ gốc/dịch theo từng ô, bên cạnh các sheet đã dịch.
- **PDF lớn:** tự phát hiện số trang; nếu PDF vượt quá 15 trang, app cho phép
  **chọn khoảng trang cần dịch** (thay vì bắt buộc dịch toàn bộ) để kiểm soát
  thời gian chờ và chi phí API — trích xuất và hiện tiến trình theo từng trang.
- **Chuyên ngành MEP:** giữ nguyên mã thiết bị (AHU-01, FCU...), tiêu chuẩn
  (TCVN, QCVN, ASHRAE, NFPA...), đơn vị kỹ thuật (kW, CFM, Pa...); dùng thuật
  ngữ MEP chuẩn ngành Cơ - Điện - Nước cho mọi cặp ngôn ngữ ở trên.
- **Excel:** dịch từng ô, giữ nguyên cấu trúc bảng, tải về file .xlsx đã dịch
  (hoặc bản song ngữ) — phù hợp để dịch bảng thống kê thiết bị (schedule).
- **PDF/Word/Ảnh:** trích xuất văn bản, dịch theo đoạn, xem song song bản
  gốc/bản dịch, tải về .txt.

## 6. Chia sẻ cho đồng nghiệp dùng — API key vẫn nằm ở phía bạn

Mặc định app bắt nhập API key thủ công (dành cho bạn tự dùng). Để đồng
nghiệp dùng chung mà **không thấy, không cần nhập key**, bạn cấu hình key ở
phía server một lần, code sẽ tự lấy và ẩn ô nhập đi.

### Bước dùng chung: tạo file secrets

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Mở file `.streamlit/secrets.toml` vừa tạo, điền:

```toml
ANTHROPIC_API_KEY = "sk-ant-key-that-cua-ban"
APP_PASSWORD = "mat-khau-noi-bo"   # tùy chọn — nên đặt nếu chia sẻ ra ngoài LAN
```

File này đã được thêm vào `.gitignore` — **không bao giờ commit lên GitHub**.
Khi `ANTHROPIC_API_KEY` đã có trong secrets, sidebar sẽ tự ẩn ô nhập key và
hiện dòng "🔑 API key đã được quản trị viên cấu hình sẵn" — đồng nghiệp chỉ
việc tải file và dịch, không đụng tới key.

`APP_PASSWORD` (tùy chọn) tạo một cổng mật khẩu đơn giản trước khi vào app —
nên dùng nếu bạn deploy ra ngoài mạng nội bộ, để tránh người lạ dùng ké
API key (tốn phí) của bạn. Bỏ trống hoặc xóa dòng này nếu chỉ dùng trong
văn phòng, không cần thiết.

### Cách A — Chia sẻ trong mạng nội bộ (LAN/wifi công ty), nhanh nhất

```bash
streamlit run mep_translator.py --server.address 0.0.0.0 --server.port 8501
```

Terminal sẽ in ra một dòng **Network URL**, dạng `http://192.168.x.x:8501`.
Gửi địa chỉ này cho đồng nghiệp **đang cùng mạng wifi/LAN** với máy bạn, họ
mở trình duyệt và dán vào là dùng được — máy bạn phải bật app suốt thời gian
họ dùng. Phù hợp dùng trong văn phòng, không cần đăng ký gì thêm.

### Cách B — Deploy lên Streamlit Community Cloud, dùng được từ xa (miễn phí)

Phù hợp nếu đồng nghiệp không cùng mạng, hoặc bạn muốn app chạy 24/7 không
cần bật máy tính cá nhân.

1. Push code (không kèm `secrets.toml` thật) lên repo GitHub của bạn — dùng
   đúng quy trình đã hướng dẫn ở phần trước (`git add . / commit / push`).
2. Vào **share.streamlit.io**, đăng nhập bằng tài khoản GitHub.
3. Bấm **New app**, chọn repo (VD: `ThanhNgo81-newdata/ideal-robot`), chọn
   nhánh `main` và file chính `mep_translator.py`.
4. Trước khi Deploy, vào mục **Advanced settings → Secrets**, dán y hệt nội
   dung file `secrets.toml` của bạn vào đó (đây là nơi lưu bí mật riêng của
   Streamlit Cloud, đã mã hóa, chỉ mình bạn thấy — không liên quan gì tới
   file trên GitHub).
5. Bấm **Deploy**. Sau vài phút, bạn nhận được một URL công khai dạng
   `https://ten-app.streamlit.app` — gửi link này cho đồng nghiệp, họ dùng
   được từ bất kỳ đâu có internet, key vẫn chỉ nằm trong Secrets của bạn.

## 7. Lưu ý

- PDF dạng scan ảnh (không có lớp văn bản) sẽ không trích xuất được bằng
  pypdf — trong trường hợp đó, chụp ảnh trang cần dịch và tải lên như ảnh.
- Nếu không cấu hình `secrets.toml`, key nhập tay chỉ lưu tạm trong phiên
  trình duyệt, không ghi ra file — mỗi lần mở lại app cần nhập lại.
- Tài liệu rất dài sẽ được chia nhỏ để dịch, có thể mất từ vài chục giây đến
  vài phút tùy độ dài.
- PDF càng nhiều trang, số lượt gọi API càng nhiều → thời gian chờ và chi phí
  càng tăng tuyến tính. Với PDF rất lớn (trăm trang), nên dùng tính năng chọn
  khoảng trang để dịch từng phần thay vì dịch một lần toàn bộ tài liệu.
- Chế độ song ngữ dùng cùng số lượt gọi API như chế độ dịch thường (chỉ đổi
  định dạng phản hồi), nhưng phản hồi dài hơn (có cả gốc lẫn dịch) nên có thể
  tốn nhiều token đầu ra hơn một chút.
