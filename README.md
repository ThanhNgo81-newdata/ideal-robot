# 🛠️ MEP Translator (Offline)

Ứng dụng dịch tài liệu kỹ thuật MEP (Cơ – Điện – Nước) hỗ trợ **DOCX, PDF, Excel, TXT**.  
Chạy hoàn toàn **offline bằng Hugging Face**, không cần API key.  
Có chế độ **song ngữ**: giữ nguyên nội dung gốc và hiển thị bản dịch ngay bên dưới hoặc bên cạnh.

---

## ✨ Tính năng
- Dịch đa ngôn ngữ (Anh, Việt, Pháp, Đức, Trung, Nhật, Hàn, Tây Ban Nha, Nga, Bồ Đào Nha, Ý…).
- Giữ nguyên cấu trúc file gốc (Word, Excel, PDF).
- Chế độ song ngữ:
  - DOCX: đoạn gốc + bản dịch ngay dưới.
  - PDF: xuất ra DOCX, block gốc + bản dịch ngay dưới.
  - Excel: chèn thêm cột kế bên chứa bản dịch.
  - TXT: ghép gốc + dịch trong cùng file.
- Giao diện web thân thiện với **Streamlit**.

---

## ⚙️ Cài đặt và chạy local

### 1. Clone repo
```bash
git clone https://github.com/<username>/<repo>.git
cd <repo>

## ⚙️ Tạo môi trường ảo

### 2. Tạo môi trường ảo (khuyến nghị)
```bash
python -m venv venv
source venv/bin/activate   # macOS/Linux
venv\Scripts\activate      # Windows

### 3. Cài dặt thư viện
```bash
pip install -r requirements.txt
