# -*- coding: utf-8 -*-
"""
MEP Translator — Ứng dụng dịch tài liệu kỹ thuật MEP (PDF lớn, Word, Excel, Ảnh)
Hỗ trợ đa ngôn ngữ (như Google Translate) + chế độ song ngữ (đối chiếu gốc/dịch).
Chạy bằng: streamlit run mep_translator.py
"""

import io
import os
import json
import base64

import streamlit as st
import anthropic
from pypdf import PdfReader
import docx
import openpyxl
from PIL import Image

# ----------------------------------------------------------------------------
# Cấu hình chung
# ----------------------------------------------------------------------------

MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 1536
CHUNK_CHARS = 3000
XLSX_BATCH_CHARS = 1800
PDF_LARGE_THRESHOLD_PAGES = 15  # từ ngưỡng này, cho phép chọn khoảng trang

# Danh sách ngôn ngữ phổ biến (giống Google Translate) — có thể gõ tay ngôn
# ngữ khác không có trong danh sách qua lựa chọn "Khác...".
LANGUAGE_PRESETS = [
    ("en", "Tiếng Anh", "English"),
    ("vi", "Tiếng Việt", "Vietnamese"),
    ("ja", "Tiếng Nhật", "Japanese"),
    ("zh-Hans", "Tiếng Trung (giản thể)", "Chinese (Simplified)"),
    ("zh-Hant", "Tiếng Trung (phồn thể)", "Chinese (Traditional)"),
    ("ko", "Tiếng Hàn", "Korean"),
    ("th", "Tiếng Thái", "Thai"),
    ("fr", "Tiếng Pháp", "French"),
    ("de", "Tiếng Đức", "German"),
    ("es", "Tiếng Tây Ban Nha", "Spanish"),
    ("pt", "Tiếng Bồ Đào Nha", "Portuguese"),
    ("it", "Tiếng Ý", "Italian"),
    ("ru", "Tiếng Nga", "Russian"),
    ("id", "Tiếng Indonesia", "Indonesian"),
    ("ms", "Tiếng Mã Lai", "Malay"),
    ("km", "Tiếng Khmer", "Khmer"),
    ("lo", "Tiếng Lào", "Lao"),
    ("hi", "Tiếng Hindi", "Hindi"),
    ("ar", "Tiếng Ả Rập", "Arabic"),
    ("other", "Khác... (tự nhập)", None),
]
VI_LABEL = {code: vi for code, vi, _ in LANGUAGE_PRESETS}
EN_LABEL = {code: en for code, _, en in LANGUAGE_PRESETS}

GLOSSARY_HINT = (
    "AHU=Air Handling Unit/Bộ xử lý không khí; FCU=Fan Coil Unit; "
    "VRV/VRF=máy điều hòa trung tâm; MCCB/ACB=Aptomat; "
    "tủ điện=electrical panel/switchboard; ống gió=duct/air duct; "
    "ống nước=pipe/piping; PCCC=fire fighting/fire protection; "
    "máy bơm=pump; chiller=máy làm lạnh nước; thông gió=ventilation; "
    "cấp thoát nước=water supply and drainage."
)

st.set_page_config(page_title="MEP Translator", page_icon="🛠️", layout="wide")


# ----------------------------------------------------------------------------
# Đọc cấu hình bí mật (API key, mật khẩu truy cập) — server-side, không lộ
# cho người dùng. Ưu tiên st.secrets (deploy Streamlit Cloud), sau đó tới
# biến môi trường (chạy local/LAN).
# ----------------------------------------------------------------------------

def get_secret(key: str, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key, default)


ADMIN_API_KEY = get_secret("ANTHROPIC_API_KEY")
APP_PASSWORD = get_secret("APP_PASSWORD")  # để trống nếu không cần cổng mật khẩu


def check_password_gate():
    if not APP_PASSWORD:
        return True
    if st.session_state.get("authed"):
        return True
    st.title("🛠️ MEP TRANSLATOR")
    st.caption("Nhập mật khẩu truy cập do quản trị viên cung cấp.")
    pw = st.text_input("Mật khẩu", type="password")
    if st.button("Vào ứng dụng"):
        if pw == APP_PASSWORD:
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("Sai mật khẩu.")
    return False


# ----------------------------------------------------------------------------
# Tiện ích dịch thuật
# ----------------------------------------------------------------------------

def resolve_lang_label(code: str, custom_name: str) -> str:
    """Trả về tên ngôn ngữ để đưa vào prompt (tiếng Anh cho model dễ hiểu)."""
    if code == "auto":
        return "ngôn ngữ nguồn (tự động nhận diện)"
    if code == "other":
        return custom_name.strip() or "ngôn ngữ do người dùng chỉ định"
    return EN_LABEL.get(code, code)


def build_system_prompt(src_label: str, tgt_label: str, bilingual: bool) -> str:
    bilingual_rule = ""
    if bilingual:
        bilingual_rule = (
            "\n- CHẾ ĐỘ SONG NGỮ: với mỗi đoạn được giao, xuất theo đúng khuôn:\n"
            "  [GỐC]\n<toàn văn đoạn gốc, giữ nguyên>\n\n[DỊCH]\n<bản dịch đoạn đó>\n\n"
            "  Không thêm gì khác ngoài khuôn này."
        )
    return (
        f"Bạn là chuyên gia dịch thuật kỹ thuật chuyên ngành MEP (Cơ - Điện - Nước: "
        f"HVAC, hệ thống điện, cấp thoát nước, phòng cháy chữa cháy) trong xây dựng.\n"
        f"Nhiệm vụ: dịch nội dung từ {src_label} sang {tgt_label}.\n"
        "Quy tắc bắt buộc:\n"
        "- Giữ nguyên, không dịch: mã hiệu thiết bị (VD: AHU-01, FCU-12, DB-3F), "
        "số hiệu tiêu chuẩn (TCVN, QCVN, ASHRAE, NFPA, ASME, SMACNA, IEC), model, "
        "mã sản phẩm, và đơn vị kỹ thuật (kW, CFM, m3/h, Pa, mmAq, °C, kVA, mm, A, V).\n"
        f"- Sử dụng thuật ngữ MEP chuẩn ngành. Tham khảo: {GLOSSARY_HINT}\n"
        "- Giữ nguyên cấu trúc định dạng, bảng biểu, số thứ tự và xuống dòng của "
        "văn bản gốc càng sát càng tốt."
        + bilingual_rule +
        "\n- Không thêm ghi chú, lời giải thích ngoài yêu cầu, không dùng markdown code fence."
    )


def get_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


def translate_text_chunk(client, system_prompt: str, chunk: str, bilingual: bool) -> str:
    instruction = (
        "Dịch đoạn văn bản kỹ thuật MEP sau đây theo đúng khuôn [GỐC]/[DỊCH] đã quy định "
        "(một phần trong tài liệu lớn hơn):\n\n" + chunk
        if bilingual else
        "Dịch đoạn văn bản kỹ thuật MEP sau đây (một phần trong tài liệu lớn "
        "hơn, không thêm tiêu đề hay bình luận):\n\n" + chunk
    )
    msg = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=system_prompt,
        messages=[{"role": "user", "content": instruction}],
    )
    return "".join(b.text for b in msg.content if b.type == "text").strip()


def translate_batch_strings(client, system_prompt: str, items: list) -> list:
    """Dịch một mảng chuỗi ngắn (dùng cho ô Excel), trả về mảng cùng độ dài."""
    prompt = (
        system_prompt
        + "\n\nDịch từng chuỗi trong mảng JSON sau đây (nội dung ô bảng tính kỹ "
          "thuật MEP). Chỉ trả về DUY NHẤT một mảng JSON hợp lệ chứa bản dịch, "
          "cùng số phần tử và đúng thứ tự, không thêm giải thích, không dùng "
          "code fence:\n\n" + json.dumps(items, ensure_ascii=False)
    )
    msg = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = "".join(b.text for b in msg.content if b.type == "text").strip().strip("`")
    if raw.lower().startswith("json"):
        raw = raw[4:].strip()
    try:
        arr = json.loads(raw)
        if isinstance(arr, list) and len(arr) == len(items):
            return [str(x) for x in arr]
    except Exception:
        pass
    return items  # fallback: giữ nguyên nếu parse lỗi, để không phá cấu trúc


def translate_image(client, system_prompt: str, tgt_label: str, image_bytes: bytes,
                     media_type: str, bilingual: bool) -> str:
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    if bilingual:
        instruction = (
            "Đây là ảnh chụp tài liệu/bản vẽ kỹ thuật MEP. Hãy nhận diện toàn bộ văn bản "
            f"trong ảnh, sau đó xuất theo đúng khuôn:\n[GỐC]\n<toàn bộ văn bản nhận diện được>\n\n"
            f"[DỊCH]\n<bản dịch sang {tgt_label}>\n\nGiữ đúng bố cục bảng biểu nếu có (dùng dấu | "
            "để phân cách cột)."
        )
    else:
        instruction = (
            f"Đây là ảnh chụp tài liệu/bản vẽ kỹ thuật MEP. Hãy nhận diện toàn "
            f"bộ văn bản trong ảnh và dịch sang {tgt_label}. Giữ đúng bố cục "
            "bảng biểu nếu có (dùng dấu | để phân cách cột). Chỉ xuất bản dịch, "
            "không giải thích."
        )
    msg = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=system_prompt,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": instruction},
            ],
        }],
    )
    return "".join(b.text for b in msg.content if b.type == "text").strip()


def chunk_text(text: str, max_chars: int = CHUNK_CHARS) -> list:
    paras = text.split("\n\n")
    chunks, cur = [], ""
    for p in paras:
        if len(cur) + len(p) + 2 > max_chars and cur:
            chunks.append(cur)
            cur = p
        else:
            cur = (cur + "\n\n" + p) if cur else p
    if cur:
        chunks.append(cur)
    final = []
    for c in chunks:
        if len(c) <= max_chars:
            final.append(c)
        else:
            final.extend(c[i:i + max_chars] for i in range(0, len(c), max_chars))
    return final


# ----------------------------------------------------------------------------
# Trích xuất nội dung từ tệp
# ----------------------------------------------------------------------------

def pdf_page_count(file_bytes: bytes) -> int:
    return len(PdfReader(io.BytesIO(file_bytes)).pages)


def extract_pdf(file_bytes: bytes, page_from: int = None, page_to: int = None, progress_cb=None) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    n = len(reader.pages)
    start = (page_from - 1) if page_from else 0
    end = page_to if page_to else n
    start = max(0, min(start, n))
    end = max(start, min(end, n))
    parts = []
    total = end - start
    for idx, i in enumerate(range(start, end), start=1):
        t = reader.pages[i].extract_text() or ""
        parts.append((f"\n\n--- Trang {i+1} ---\n\n" if idx > 1 else "") + t)
        if progress_cb and total > 0:
            progress_cb(idx / total, i + 1)
    text = "".join(parts).strip()
    if not text:
        raise ValueError("Không tìm thấy văn bản trong PDF (có thể là bản scan ảnh). Hãy dùng chế độ ảnh.")
    return text


def extract_docx(file_bytes: bytes) -> str:
    d = docx.Document(io.BytesIO(file_bytes))
    paras = [p.text for p in d.paragraphs if p.text.strip()]
    for table in d.tables:
        for row in table.rows:
            paras.append(" | ".join(c.text for c in row.cells))
    text = "\n\n".join(paras).strip()
    if not text:
        raise ValueError("Tài liệu Word không có nội dung văn bản.")
    return text


def extract_xlsx(file_bytes: bytes):
    """Trả về (workbook, list_of_cell_refs, preview_text)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    refs = []
    preview_lines = []
    for ws in wb.worksheets:
        preview_lines.append(f"\n=== SHEET: {ws.title} ===")
        for row in ws.iter_rows():
            row_vals = []
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    refs.append({"sheet": ws.title, "coord": cell.coordinate, "text": cell.value})
                row_vals.append("" if cell.value is None else str(cell.value))
            if any(v.strip() for v in row_vals):
                preview_lines.append(" | ".join(row_vals))
    if not refs:
        raise ValueError("Không tìm thấy nội dung dạng chữ trong bảng tính.")
    return wb, refs, "\n".join(preview_lines).strip()


def build_bilingual_xlsx(wb, refs) -> bytes:
    """Thêm sheet đối chiếu Gốc/Dịch, đồng thời ghi đè workbook bằng bản dịch."""
    ws2 = wb.create_sheet("Song ngữ (đối chiếu)")
    ws2.append(["Sheet gốc", "Ô", "Nội dung gốc", "Bản dịch"])
    for r in refs:
        ws2.append([r["sheet"], r["coord"], r["text"], r.get("translated", "")])
    for r in refs:
        wb[r["sheet"]][r["coord"]] = r.get("translated", r["text"])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ----------------------------------------------------------------------------
# Giao diện Streamlit
# ----------------------------------------------------------------------------

st.markdown(
    """
    <style>
    .stApp { background-color: #06192b; }
    .block-container { padding-top: 2rem; }
    h1, h2, h3 { font-family: 'IBM Plex Mono', monospace; color: #eaf2f5; }
    .stMarkdown, p, label, .stCaption { color: #eaf2f5; }
    </style>
    """,
    unsafe_allow_html=True,
)

if not check_password_gate():
    st.stop()

st.title("🛠️ MEP TRANSLATOR")
st.caption("Dịch tài liệu kỹ thuật MEP (Cơ – Điện – Nước) từ PDF lớn, Word, Excel, Ảnh — đa ngôn ngữ, hỗ trợ song ngữ.")

with st.sidebar:
    st.header("Cấu hình")
    if ADMIN_API_KEY:
        api_key = ADMIN_API_KEY
        st.success("🔑 API key đã được quản trị viên cấu hình sẵn.")
    else:
        api_key = st.text_input("Anthropic API key", type="password", help="Dạng sk-ant-... — lấy tại console.anthropic.com")
    st.markdown("---")

    src_codes = ["auto"] + [c for c, _, _ in LANGUAGE_PRESETS]
    src = st.selectbox("Ngôn ngữ nguồn", src_codes,
                        format_func=lambda k: "Tự động nhận diện" if k == "auto" else VI_LABEL[k])
    src_custom = ""
    if src == "other":
        src_custom = st.text_input("Nhập tên ngôn ngữ nguồn", placeholder="VD: Tiếng Miến Điện / Burmese")

    tgt_codes = [c for c, _, _ in LANGUAGE_PRESETS]
    tgt = st.selectbox("Ngôn ngữ đích", tgt_codes, index=1, format_func=lambda k: VI_LABEL[k])
    tgt_custom = ""
    if tgt == "other":
        tgt_custom = st.text_input("Nhập tên ngôn ngữ đích", placeholder="VD: Tiếng Miến Điện / Burmese")

    st.markdown("---")
    bilingual = st.toggle("📖 Chế độ song ngữ (đối chiếu Gốc/Dịch)", value=False,
                           help="Kết quả hiển thị cả đoạn gốc và bản dịch xen kẽ, thay vì chỉ bản dịch.")
    st.markdown("---")
    st.caption("Chuyên ngành: tối ưu cho kỹ sư MEP — giữ nguyên mã thiết bị, tiêu chuẩn (TCVN/QCVN/ASHRAE/NFPA) và đơn vị kỹ thuật.")

uploaded = st.file_uploader(
    "Tải lên tệp cần dịch (PDF, DOCX, XLSX/XLS, JPG, PNG) — hỗ trợ PDF nhiều trang",
    type=["pdf", "docx", "xlsx", "xls", "jpg", "jpeg", "png", "webp"],
)

if "result" not in st.session_state:
    st.session_state.result = None

if uploaded is not None:
    kind = uploaded.name.lower().rsplit(".", 1)[-1]
    kind = "xlsx" if kind in ("xlsx", "xls") else ("image" if kind in ("jpg", "jpeg", "png", "webp") else kind)
    file_bytes = uploaded.getvalue()

    st.write(f"**Tệp:** {uploaded.name} · **Loại:** {kind.upper()} · **Kích thước:** {len(file_bytes)/1024:.0f} KB")

    page_from, page_to = None, None
    if kind == "pdf":
        try:
            n_pages = pdf_page_count(file_bytes)
            st.write(f"**Số trang:** {n_pages}")
            if n_pages > PDF_LARGE_THRESHOLD_PAGES:
                st.info(
                    f"PDF này có {n_pages} trang — để kiểm soát thời gian/chi phí, bạn có thể "
                    "chọn khoảng trang cần dịch thay vì dịch toàn bộ."
                )
                pr = st.slider("Khoảng trang cần dịch", 1, n_pages, (1, min(n_pages, 30)))
                page_from, page_to = pr[0], pr[1]
        except Exception:
            n_pages = None

    if st.button("🌐 Dịch tài liệu", type="primary", disabled=not api_key):
        if not api_key:
            st.error("Vui lòng nhập Anthropic API key ở thanh bên trái.")
        elif src == "other" and not src_custom.strip():
            st.error("Vui lòng nhập tên ngôn ngữ nguồn.")
        elif tgt == "other" and not tgt_custom.strip():
            st.error("Vui lòng nhập tên ngôn ngữ đích.")
        else:
            try:
                client = get_client(api_key)
                src_label = resolve_lang_label(src, src_custom)
                tgt_label = resolve_lang_label(tgt, tgt_custom)
                system_prompt = build_system_prompt(src_label, tgt_label, bilingual)

                progress = st.progress(0, text="Đang xử lý...")

                if kind == "pdf":
                    def _extract_progress(frac, page_no):
                        progress.progress(min(int(frac * 15), 15), text=f"Đang trích xuất trang {page_no}...")
                    original = extract_pdf(file_bytes, page_from, page_to, progress_cb=_extract_progress)
                    chunks = chunk_text(original)
                    out = []
                    for i, c in enumerate(chunks):
                        pct = 15 + int((i / len(chunks)) * 85)
                        progress.progress(pct, text=f"Đang dịch đoạn {i+1}/{len(chunks)}")
                        out.append(translate_text_chunk(client, system_prompt, c, bilingual))
                    translated = "\n\n".join(out)
                    st.session_state.result = {"kind": "text", "original": original, "translated": translated, "name": uploaded.name}

                elif kind == "docx":
                    original = extract_docx(file_bytes)
                    chunks = chunk_text(original)
                    out = []
                    for i, c in enumerate(chunks):
                        progress.progress(int((i / len(chunks)) * 100), text=f"Đang dịch đoạn {i+1}/{len(chunks)}")
                        out.append(translate_text_chunk(client, system_prompt, c, bilingual))
                    translated = "\n\n".join(out)
                    st.session_state.result = {"kind": "text", "original": original, "translated": translated, "name": uploaded.name}

                elif kind == "xlsx":
                    wb, refs, preview = extract_xlsx(file_bytes)
                    batches, cur, cur_len = [], [], 0
                    for r in refs:
                        l = len(r["text"]) + 4
                        if cur_len + l > XLSX_BATCH_CHARS and cur:
                            batches.append(cur)
                            cur, cur_len = [], 0
                        cur.append(r)
                        cur_len += l
                    if cur:
                        batches.append(cur)

                    for bi, batch in enumerate(batches):
                        progress.progress(int((bi / len(batches)) * 100), text=f"Đang dịch bảng tính — lô {bi+1}/{len(batches)}")
                        texts = [r["text"] for r in batch]
                        translated_vals = translate_batch_strings(client, system_prompt, texts)
                        for r, tval in zip(batch, translated_vals):
                            r["translated"] = tval

                    if bilingual:
                        workbook_bytes = build_bilingual_xlsx(wb, refs)
                        translated_preview = "\n".join(
                            f"[{r['sheet']}!{r['coord']}]\nGỐC: {r['text']}\nDỊCH: {r.get('translated','')}\n" for r in refs
                        )
                    else:
                        for r in refs:
                            wb[r["sheet"]][r["coord"]] = r.get("translated", r["text"])
                        out_buf = io.BytesIO()
                        wb.save(out_buf)
                        workbook_bytes = out_buf.getvalue()
                        translated_preview = "\n".join(
                            f"{r['sheet']}!{r['coord']}: {r.get('translated','')}" for r in refs
                        )

                    st.session_state.result = {
                        "kind": "xlsx", "original": preview, "translated": translated_preview,
                        "workbook_bytes": workbook_bytes, "name": uploaded.name,
                    }

                elif kind == "image":
                    media_type = uploaded.type or "image/png"
                    translated = translate_image(client, system_prompt, tgt_label, file_bytes, media_type, bilingual)
                    progress.progress(100, text="Hoàn tất")
                    st.session_state.result = {"kind": "image", "original": "(nội dung ảnh)", "translated": translated, "name": uploaded.name}

                progress.progress(100, text="Hoàn tất")
                st.success("Đã dịch xong.")
            except Exception as e:
                st.error(f"Lỗi khi dịch: {e}")

if st.session_state.result:
    res = st.session_state.result
    tab1, tab2 = st.tabs(["📄 Văn bản gốc", "🌐 Bản dịch" + (" (song ngữ)" if bilingual else "")])
    with tab1:
        st.text_area("Gốc", res["original"], height=400, label_visibility="collapsed")
    with tab2:
        st.text_area("Dịch", res["translated"], height=400, label_visibility="collapsed")

        base_name = res["name"].rsplit(".", 1)[0]
        suffix = "_bilingual" if bilingual else "_translated"
        st.download_button("⬇️ Tải bản dịch (.txt)", res["translated"], file_name=f"{base_name}{suffix}.txt")
        if res["kind"] == "xlsx" and "workbook_bytes" in res:
            st.download_button(
                "⬇️ Tải bảng tính đã dịch (.xlsx)",
                res["workbook_bytes"],
                file_name=f"{base_name}{suffix}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
