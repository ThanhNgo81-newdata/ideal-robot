# -*- coding: utf-8 -*-
"""
MEP Translator — Ứng dụng dịch tài liệu kỹ thuật MEP (PDF, Word, Excel, Ảnh)
Chạy bằng: streamlit run mep_translator.py
"""

import io
import json
import base64
import copy

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
MAX_TOKENS = 1024
CHUNK_CHARS = 3000

LANG_NAMES_VI = {"en": "Tiếng Anh", "vi": "Tiếng Việt", "zh": "Tiếng Trung"}
LANG_NAMES_EN = {"en": "English", "vi": "Vietnamese", "zh": "Chinese"}

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
# Tiện ích dịch thuật
# ----------------------------------------------------------------------------

def build_system_prompt(src_label: str, tgt_label: str) -> str:
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
        "văn bản gốc càng sát càng tốt.\n"
        "- CHỈ xuất ra nội dung đã dịch. Không thêm ghi chú, lời giải thích, không "
        "lặp lại văn bản gốc, không dùng markdown code fence."
    )


def get_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


def translate_text_chunk(client, system_prompt: str, chunk: str) -> str:
    msg = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=system_prompt,
        messages=[{
            "role": "user",
            "content": (
                "Dịch đoạn văn bản kỹ thuật MEP sau đây (một phần trong tài liệu lớn "
                "hơn, không thêm tiêu đề hay bình luận):\n\n" + chunk
            ),
        }],
    )
    return "".join(b.text for b in msg.content if b.type == "text").strip()


def translate_batch_strings(client, system_prompt: str, tgt_label: str, items: list) -> list:
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
    raw = "".join(b.text for b in msg.content if b.type == "text").strip()
    raw = raw.strip("`")
    if raw.lower().startswith("json"):
        raw = raw[4:].strip()
    try:
        arr = json.loads(raw)
        if isinstance(arr, list) and len(arr) == len(items):
            return [str(x) for x in arr]
    except Exception:
        pass
    return items  # fallback: giữ nguyên nếu parse lỗi, để không phá cấu trúc


def translate_image(client, system_prompt: str, tgt_label: str, image_bytes: bytes, media_type: str) -> str:
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    msg = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=system_prompt,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": (
                    f"Đây là ảnh chụp tài liệu/bản vẽ kỹ thuật MEP. Hãy nhận diện toàn "
                    f"bộ văn bản trong ảnh và dịch sang {tgt_label}. Giữ đúng bố cục "
                    "bảng biểu nếu có (dùng dấu | để phân cách cột). Chỉ xuất bản dịch, "
                    "không giải thích."
                )},
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

def extract_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    parts = []
    for i, page in enumerate(reader.pages, start=1):
        t = page.extract_text() or ""
        parts.append((f"\n\n--- Trang {i} ---\n\n" if i > 1 else "") + t)
    text = "".join(parts).strip()
    if not text:
        raise ValueError("Không tìm thấy văn bản trong PDF (có thể là bản scan ảnh). Hãy dùng chế độ ảnh.")
    return text


def extract_docx(file_bytes: bytes) -> str:
    d = docx.Document(io.BytesIO(file_bytes))
    paras = [p.text for p in d.paragraphs if p.text.strip()]
    for table in d.tables:
        for row in table.rows:
            cells = [c.text for c in row.cells]
            paras.append(" | ".join(cells))
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

st.title("🛠️ MEP TRANSLATOR")
st.caption("Dịch tài liệu kỹ thuật MEP (Cơ – Điện – Nước) từ PDF, Word, Excel, Ảnh — chạy hoàn toàn trên máy bạn.")

with st.sidebar:
    st.header("Cấu hình")
    api_key = st.text_input("Anthropic API key", type="password", help="Dạng sk-ant-... — lấy tại console.anthropic.com")
    st.markdown("---")
    src = st.selectbox("Ngôn ngữ nguồn", ["auto", "en", "vi", "zh"],
                        format_func=lambda k: "Tự động nhận diện" if k == "auto" else LANG_NAMES_VI[k])
    tgt = st.selectbox("Ngôn ngữ đích", ["vi", "en", "zh"], format_func=lambda k: LANG_NAMES_VI[k])
    st.markdown("---")
    st.caption("Chuyên ngành: tối ưu cho kỹ sư MEP — giữ nguyên mã thiết bị, tiêu chuẩn (TCVN/QCVN/ASHRAE/NFPA) và đơn vị kỹ thuật.")

uploaded = st.file_uploader(
    "Tải lên tệp cần dịch (PDF, DOCX, XLSX/XLS, JPG, PNG)",
    type=["pdf", "docx", "xlsx", "xls", "jpg", "jpeg", "png", "webp"],
)

if "result" not in st.session_state:
    st.session_state.result = None

if uploaded is not None:
    kind = uploaded.name.lower().rsplit(".", 1)[-1]
    kind = "xlsx" if kind in ("xlsx", "xls") else ("image" if kind in ("jpg", "jpeg", "png", "webp") else kind)
    file_bytes = uploaded.getvalue()

    st.write(f"**Tệp:** {uploaded.name} · **Loại:** {kind.upper()} · **Kích thước:** {len(file_bytes)/1024:.0f} KB")

    col1, col2 = st.columns(2)

    if st.button("🌐 Dịch tài liệu", type="primary", disabled=not api_key):
        if not api_key:
            st.error("Vui lòng nhập Anthropic API key ở thanh bên trái.")
        else:
            try:
                client = get_client(api_key)
                src_label = "ngôn ngữ nguồn (tự động nhận diện)" if src == "auto" else LANG_NAMES_EN[src]
                tgt_label = LANG_NAMES_EN[tgt]
                system_prompt = build_system_prompt(src_label, tgt_label)

                progress = st.progress(0, text="Đang xử lý...")

                if kind == "pdf":
                    original = extract_pdf(file_bytes)
                    chunks = chunk_text(original)
                    out = []
                    for i, c in enumerate(chunks):
                        progress.progress(int((i / len(chunks)) * 100), text=f"Đang dịch đoạn {i+1}/{len(chunks)}")
                        out.append(translate_text_chunk(client, system_prompt, c))
                    translated = "\n\n".join(out)
                    st.session_state.result = {"kind": "text", "original": original, "translated": translated, "name": uploaded.name}

                elif kind == "docx":
                    original = extract_docx(file_bytes)
                    chunks = chunk_text(original)
                    out = []
                    for i, c in enumerate(chunks):
                        progress.progress(int((i / len(chunks)) * 100), text=f"Đang dịch đoạn {i+1}/{len(chunks)}")
                        out.append(translate_text_chunk(client, system_prompt, c))
                    translated = "\n\n".join(out)
                    st.session_state.result = {"kind": "text", "original": original, "translated": translated, "name": uploaded.name}

                elif kind == "xlsx":
                    wb, refs, preview = extract_xlsx(file_bytes)
                    batches, cur, cur_len = [], [], 0
                    for r in refs:
                        l = len(r["text"]) + 4
                        if cur_len + l > 1800 and cur:
                            batches.append(cur)
                            cur, cur_len = [], 0
                        cur.append(r)
                        cur_len += l
                    if cur:
                        batches.append(cur)

                    for bi, batch in enumerate(batches):
                        progress.progress(int((bi / len(batches)) * 100), text=f"Đang dịch bảng tính — lô {bi+1}/{len(batches)}")
                        texts = [r["text"] for r in batch]
                        translated_vals = translate_batch_strings(client, system_prompt, tgt_label, texts)
                        for r, tval in zip(batch, translated_vals):
                            r["translated"] = tval

                    # Ghi lại vào workbook
                    for r in refs:
                        wb[r["sheet"]][r["coord"]] = r.get("translated", r["text"])

                    out_buf = io.BytesIO()
                    wb.save(out_buf)
                    translated_preview = "\n".join(
                        f"{r['sheet']}!{r['coord']}: {r.get('translated','')}" for r in refs
                    )
                    st.session_state.result = {
                        "kind": "xlsx", "original": preview, "translated": translated_preview,
                        "workbook_bytes": out_buf.getvalue(), "name": uploaded.name,
                    }

                elif kind == "image":
                    media_type = uploaded.type or "image/png"
                    translated = translate_image(client, system_prompt, tgt_label, file_bytes, media_type)
                    progress.progress(100, text="Hoàn tất")
                    st.session_state.result = {"kind": "image", "original": "(nội dung ảnh)", "translated": translated, "name": uploaded.name}

                progress.progress(100, text="Hoàn tất")
                st.success("Đã dịch xong.")
            except Exception as e:
                st.error(f"Lỗi khi dịch: {e}")

if st.session_state.result:
    res = st.session_state.result
    tab1, tab2 = st.tabs(["📄 Văn bản gốc", "🌐 Bản dịch"])
    with tab1:
        st.text_area("Gốc", res["original"], height=400, label_visibility="collapsed")
    with tab2:
        st.text_area("Dịch", res["translated"], height=400, label_visibility="collapsed")

        base_name = res["name"].rsplit(".", 1)[0]
        st.download_button("⬇️ Tải bản dịch (.txt)", res["translated"], file_name=f"{base_name}_translated.txt")
        if res["kind"] == "xlsx" and "workbook_bytes" in res:
            st.download_button(
                "⬇️ Tải bảng tính đã dịch (.xlsx)",
                res["workbook_bytes"],
                file_name=f"{base_name}_translated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
