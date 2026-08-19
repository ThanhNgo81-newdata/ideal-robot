# -*- coding: utf-8 -*-
"""
MEP Translator v2 — Dịch tài liệu kỹ thuật MEP giữ nguyên cấu trúc file gốc
(PDF lớn, Word, Excel, Ảnh/ảnh chụp màn hình) — đa ngôn ngữ, có chế độ song ngữ.

Chạy: streamlit run mep_translator.py
"""

import io
import json
import base64
import copy

import streamlit as st
import anthropic
import pymupdf as fitz
import docx
from docx.text.paragraph import Paragraph
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter
from openpyxl.cell.cell import MergedCell

# ----------------------------------------------------------------------------
# Cấu hình chung
# ----------------------------------------------------------------------------

MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 1500
CHUNK_CHARS = 3000
SEP = "\n@@@BLOCK@@@\n"           # dấu phân cách khi dịch nhiều đoạn trong 1 lần gọi
PDF_PAGE_WARN_THRESHOLD = 25      # PDF trên ngưỡng này sẽ gợi ý chọn khoảng trang

LANGUAGES_VI = {
    "en": "Tiếng Anh", "vi": "Tiếng Việt", "zh": "Tiếng Trung", "ja": "Tiếng Nhật",
    "ko": "Tiếng Hàn", "fr": "Tiếng Pháp", "de": "Tiếng Đức", "th": "Tiếng Thái",
    "es": "Tiếng Tây Ban Nha", "ru": "Tiếng Nga", "id": "Tiếng Indonesia",
    "ms": "Tiếng Mã Lai", "km": "Tiếng Khmer", "lo": "Tiếng Lào",
    "hi": "Tiếng Hindi", "ar": "Tiếng Ả Rập", "pt": "Tiếng Bồ Đào Nha", "it": "Tiếng Ý",
}
LANGUAGES_EN = {
    "en": "English", "vi": "Vietnamese", "zh": "Chinese", "ja": "Japanese",
    "ko": "Korean", "fr": "French", "de": "German", "th": "Thai",
    "es": "Spanish", "ru": "Russian", "id": "Indonesian", "ms": "Malay",
    "km": "Khmer", "lo": "Lao", "hi": "Hindi", "ar": "Arabic",
    "pt": "Portuguese", "it": "Italian",
}

GLOSSARY_HINT = (
    "AHU=Air Handling Unit/Bo xu ly khong khi; FCU=Fan Coil Unit; "
    "VRV/VRF=may dieu hoa trung tam; MCCB/ACB=Aptomat; "
    "tu dien=electrical panel/switchboard; ong gio=duct/air duct; "
    "ong nuoc=pipe/piping; PCCC=fire fighting/fire protection; "
    "may bom=pump; chiller=may lam lanh nuoc; thong gio=ventilation; "
    "cap thoat nuoc=water supply and drainage."
)

st.set_page_config(page_title="MEP Translator", page_icon="🛠️", layout="wide")


# ----------------------------------------------------------------------------
# Dịch thuật (Anthropic API)
# ----------------------------------------------------------------------------

def get_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


def build_system_prompt(src_label: str, tgt_label: str) -> str:
    return (
        f"Ban la chuyen gia dich thuat ky thuat chuyen nganh MEP (Co - Dien - Nuoc: "
        f"HVAC, he thong dien, cap thoat nuoc, phong chay chua chay) trong xay dung.\n"
        f"Nhiem vu: dich noi dung tu {src_label} sang {tgt_label}.\n"
        "Quy tac bat buoc:\n"
        "- Giu nguyen, KHONG dich: ma hieu thiet bi (VD: AHU-01, FCU-12, DB-3F), "
        "so hieu tieu chuan (TCVN, QCVN, ASHRAE, NFPA, ASME, SMACNA, IEC, JIS), model, "
        "ma san pham, va don vi ky thuat (kW, CFM, m3/h, Pa, mmAq, C, kVA, mm, A, V).\n"
        f"- Su dung thuat ngu MEP chuan nganh. Tham khao: {GLOSSARY_HINT}\n"
        "- Giu nguyen cau truc dinh dang, bang bieu, so thu tu va xuong dong cua "
        "van ban goc cang sat cang tot.\n"
        "- CHI xuat ra noi dung da dich. Khong them ghi chu, loi giai thich, khong "
        "lap lai van ban goc, khong dung markdown code fence."
    )


def translate_text(client, system_prompt: str, text: str) -> str:
    msg = client.messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=system_prompt,
        messages=[{"role": "user", "content":
            "Dich doan van ban ky thuat MEP sau (mot phan trong tai lieu lon hon, "
            "khong them tieu de hay binh luan):\n\n" + text}],
    )
    return "".join(b.text for b in msg.content if b.type == "text").strip()


def translate_blocks_batch(client, system_prompt: str, blocks: list) -> list:
    """Dịch nhiều đoạn văn bản trong 1 lần gọi API, dùng dấu phân cách riêng.
    Nếu số lượng đoạn trả về không khớp, sẽ dịch lại từng đoạn lẻ để đảm bảo an toàn."""
    if not blocks:
        return []
    joined = SEP.join(blocks)
    prompt = (
        system_prompt
        + f"\n\nDich tung doan van ban duoi day. CAC DOAN duoc ngan cach boi dong "
          f"'{SEP.strip()}' — hay GIU NGUYEN dung dau phan cach nay giua cac ban dich, "
          f"tra ve dung so luong doan ({len(blocks)} doan), dung thu tu, khong them "
          "giai thich:\n\n" + joined
    )
    msg = client.messages.create(model=MODEL, max_tokens=MAX_TOKENS * 2,
                                  messages=[{"role": "user", "content": prompt}])
    raw = "".join(b.text for b in msg.content if b.type == "text").strip()
    parts = [p.strip() for p in raw.split(SEP.strip())]
    if len(parts) == len(blocks):
        return parts
    return [translate_text(client, system_prompt, b) for b in blocks]


def translate_batch_strings(client, system_prompt: str, items: list) -> list:
    """Dịch mảng chuỗi ngắn (dùng cho ô Excel) qua JSON."""
    if not items:
        return []
    prompt = (
        system_prompt
        + "\n\nDich tung chuoi trong mang JSON sau (noi dung o bang tinh ky thuat "
          "MEP). CHI tra ve DUY NHAT mot mang JSON hop le chua ban dich, cung so "
          "phan tu va dung thu tu, khong giai thich, khong dung code fence:\n\n"
          + json.dumps(items, ensure_ascii=False)
    )
    msg = client.messages.create(model=MODEL, max_tokens=MAX_TOKENS,
                                  messages=[{"role": "user", "content": prompt}])
    raw = "".join(b.text for b in msg.content if b.type == "text").strip().strip("`")
    if raw.lower().startswith("json"):
        raw = raw[4:].strip()
    try:
        arr = json.loads(raw)
        if isinstance(arr, list) and len(arr) == len(items):
            return [str(x) for x in arr]
    except Exception:
        pass
    return items


def translate_image(client, system_prompt: str, tgt_label: str, image_bytes: bytes,
                     media_type: str, bilingual: bool) -> str:
    instr = (
        f"Day la anh chup tai lieu/ban ve/man hinh ky thuat MEP. Nhan dien toan bo "
        f"van ban trong anh va dich sang {tgt_label}. Giu dung bo cuc bang bieu neu "
        "co (dung dau | de phan cach cot)."
    )
    if bilingual:
        instr += (
            " Trinh bay SONG NGU: voi moi dong/khoi noi dung, hien dong goc truoc, "
            "ngay ben duoi la dong da dich trong ngoac [ ]."
        )
    instr += " Chi xuat noi dung da xu ly, khong giai thich them."
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    msg = client.messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=system_prompt,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
            {"type": "text", "text": instr},
        ]}],
    )
    return "".join(b.text for b in msg.content if b.type == "text").strip()


# ----------------------------------------------------------------------------
# WORD (.docx) — dịch tại chỗ, giữ nguyên style/heading/bảng/bullet
# ----------------------------------------------------------------------------

def insert_paragraph_after(paragraph: Paragraph, text: str, italic=True):
    """Chèn 1 đoạn mới ngay sau paragraph, sao chép định dạng — dùng cho bản song ngữ."""
    new_p = copy.deepcopy(paragraph._p)
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    for run in list(new_para.runs):
        run.text = ""
    if new_para.runs:
        new_para.runs[0].text = text
        new_para.runs[0].italic = italic
        for r in new_para.runs[1:]:
            r.text = ""
    else:
        r = new_para.add_run(text)
        r.italic = italic
    return new_para


def translate_docx(client, system_prompt: str, file_bytes: bytes, bilingual: bool, progress_cb=None):
    d = docx.Document(io.BytesIO(file_bytes))

    targets = []
    for p in d.paragraphs:
        if p.text.strip():
            targets.append(p)
    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip():
                        targets.append(p)

    original_preview = "\n\n".join(t.text for t in targets)

    BATCH = 12
    translated_all = []
    total = max(1, (len(targets) + BATCH - 1) // BATCH)
    for i in range(0, len(targets), BATCH):
        batch = targets[i:i + BATCH]
        texts = [t.text for t in batch]
        translated_all.extend(translate_blocks_batch(client, system_prompt, texts))
        if progress_cb:
            progress_cb(min(100, int(((i // BATCH) + 1) / total * 100)))

    for para, tvalue in reversed(list(zip(targets, translated_all))):
        if bilingual:
            insert_paragraph_after(para, tvalue, italic=True)
        else:
            runs = para.runs
            if runs:
                runs[0].text = tvalue
                for r in runs[1:]:
                    r.text = ""
            else:
                para.add_run(tvalue)

    out = io.BytesIO()
    d.save(out)
    return original_preview, "\n\n".join(translated_all), out.getvalue()


# ----------------------------------------------------------------------------
# EXCEL (.xlsx) — dịch tại chỗ, giữ nguyên sheet/style/merge/formula
# ----------------------------------------------------------------------------

def insert_column_with_merge_fix(ws, insert_idx: int):
    """Chèn 1 cột tại vị trí insert_idx (1-based), tự giãn các vùng merge bị
    ảnh hưởng để không vỡ layout (ô tiêu đề merge nhiều cột, v.v.)."""
    merges = list(ws.merged_cells.ranges)
    for m in merges:
        ws.unmerge_cells(str(m))
    ws.insert_cols(insert_idx)
    for m in merges:
        min_col, min_row, max_col, max_row = range_boundaries(str(m))
        if min_col >= insert_idx:
            min_col += 1
        if max_col >= insert_idx:
            max_col += 1
        ws.merge_cells(f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}")


def copy_cell_style(src_cell, dst_cell):
    dst_cell.font = copy.copy(src_cell.font)
    dst_cell.fill = copy.copy(src_cell.fill)
    dst_cell.border = copy.copy(src_cell.border)
    dst_cell.alignment = copy.copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format


def translate_xlsx(client, system_prompt: str, file_bytes: bytes, bilingual: bool, progress_cb=None):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    refs = []
    preview_lines = []
    for ws in wb.worksheets:
        preview_lines.append(f"\n=== SHEET: {ws.title} ===")
        for row in ws.iter_rows():
            row_vals = []
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    refs.append({"sheet": ws.title, "coord": cell.coordinate, "text": cell.value,
                                 "row": cell.row, "col": cell.column})
                row_vals.append("" if cell.value is None else str(cell.value))
            if any(v.strip() for v in row_vals):
                preview_lines.append(" | ".join(row_vals))
    if not refs:
        raise ValueError("Khong tim thay noi dung dang chu trong bang tinh.")

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
        texts = [r["text"] for r in batch]
        translated_vals = translate_batch_strings(client, system_prompt, texts)
        for r, tval in zip(batch, translated_vals):
            r["translated"] = tval
        if progress_cb:
            progress_cb(int((bi + 1) / len(batches) * 100))

    if bilingual:
        # Song ngữ theo kiểu "cột kề cột" ngay trong sheet gốc: với mỗi cột có
        # chữ, chèn 1 cột dịch ngay bên phải, giữ nguyên cột gốc — xử lý từ
        # cột phải nhất về trái để không lệch chỉ số cột khi chèn.
        # Riêng các ô tiêu đề bị merge rộng (banner nhiều cột) không chèn cột
        # được (ô bên cạnh vẫn thuộc vùng merge) — với các ô này, gộp bản dịch
        # ngay trong cùng ô (gốc / dịch).
        for ws in wb.worksheets:
            sheet_refs = [r for r in refs if r["sheet"] == ws.title]
            if not sheet_refs:
                continue

            wide_merge_coords = set()
            for m in ws.merged_cells.ranges:
                if m.max_col > m.min_col:  # merge trải rộng nhiều cột
                    wide_merge_coords.add((m.min_row, m.min_col))

            normal_refs, wide_refs = [], []
            for r in sheet_refs:
                (wide_refs if (r["row"], r["col"]) in wide_merge_coords else normal_refs).append(r)

            # Ô tiêu đề merge rộng: gộp gốc/dịch trong cùng ô
            for r in wide_refs:
                cell = ws.cell(row=r["row"], column=r["col"])
                cell.value = f"{r['text']}  /  {r.get('translated', '')}"

            # Ô dữ liệu thường: chèn cột dịch kề bên
            cols_with_text = sorted(set(r["col"] for r in normal_refs), reverse=True)
            for col_idx in cols_with_text:
                insert_column_with_merge_fix(ws, col_idx + 1)
                src_letter = get_column_letter(col_idx)
                new_letter = get_column_letter(col_idx + 1)
                if src_letter in ws.column_dimensions:
                    ws.column_dimensions[new_letter].width = ws.column_dimensions[src_letter].width
                for r in normal_refs:
                    if r["col"] == col_idx:
                        dst_cell = ws.cell(row=r["row"], column=col_idx + 1)
                        if isinstance(dst_cell, MergedCell):
                            # an toàn dự phòng: nếu vẫn rơi vào vùng merge, gộp
                            # vào ô gốc thay vì ghi đè (tránh crash)
                            src_cell = ws.cell(row=r["row"], column=col_idx)
                            src_cell.value = f"{r['text']}  /  {r.get('translated', '')}"
                            continue
                        src_cell = ws.cell(row=r["row"], column=col_idx)
                        dst_cell.value = r.get("translated", "")
                        copy_cell_style(src_cell, dst_cell)
    else:
        for r in refs:
            wb[r["sheet"]][r["coord"]] = r.get("translated", r["text"])

    out_buf = io.BytesIO()
    wb.save(out_buf)
    original_preview = "\n".join(preview_lines).strip()
    translated_preview = "\n".join(f"{r['sheet']}!{r['coord']}: {r.get('translated','')}" for r in refs)
    return original_preview, translated_preview, out_buf.getvalue()


# ----------------------------------------------------------------------------
# PDF (kể cả PDF lớn) — trích xuất theo block/trang, xuất ra DOCX giữ đúng
# thứ tự đọc + ngắt trang gốc (không ghi đè trực tiếp lên PDF để tránh lỗi
# hiển thị dấu tiếng Việt / chữ Hán / chữ Nhật do thiếu font nhúng sẵn).
# ----------------------------------------------------------------------------

def get_pdf_page_count(file_bytes: bytes) -> int:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    n = doc.page_count
    doc.close()
    return n


def extract_pdf_blocks(file_bytes: bytes, page_from: int, page_to: int):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pages_blocks = []
    for pno in range(page_from - 1, page_to):
        page = doc.load_page(pno)
        blocks = page.get_text("blocks")
        blocks = sorted(blocks, key=lambda b: (round(b[1], 1), round(b[0], 1)))
        texts = [b[4].strip() for b in blocks if b[4] and b[4].strip()]
        if texts:
            pages_blocks.append((pno + 1, texts))
    doc.close()
    if not pages_blocks:
        raise ValueError("Khong trich xuat duoc van ban trong khoang trang da chon "
                          "(co the la PDF dang scan anh — hay chup man hinh trang do "
                          "va dung che do Anh thay the).")
    return pages_blocks


def translate_pdf_to_docx(client, system_prompt: str, file_bytes: bytes,
                           page_from: int, page_to: int, bilingual: bool, progress_cb=None):
    pages_blocks = extract_pdf_blocks(file_bytes, page_from, page_to)

    out_doc = docx.Document()
    out_doc.add_heading("Bản dịch tài liệu PDF (MEP Translator)", level=1)

    original_parts, translated_parts = [], []
    total_pages = len(pages_blocks)
    for idx, (pno, texts) in enumerate(pages_blocks):
        out_doc.add_heading(f"Trang {pno}", level=2)
        translated_texts = translate_blocks_batch(client, system_prompt, texts)
        for orig, tval in zip(texts, translated_texts):
            if bilingual:
                p1 = out_doc.add_paragraph(orig)
                if p1.runs:
                    p1.runs[0].italic = True
                out_doc.add_paragraph(tval)
            else:
                out_doc.add_paragraph(tval)
        original_parts.append(f"--- Trang {pno} ---\n" + "\n\n".join(texts))
        translated_parts.append(f"--- Trang {pno} ---\n" + "\n\n".join(translated_texts))
        if progress_cb:
            progress_cb(int((idx + 1) / total_pages * 100))
        if idx < total_pages - 1:
            out_doc.add_page_break()

    out = io.BytesIO()
    out_doc.save(out)
    return "\n\n".join(original_parts), "\n\n".join(translated_parts), out.getvalue()


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
st.caption("Dịch tài liệu kỹ thuật MEP (Cơ – Điện – Nước) — PDF lớn, Word, Excel, Ảnh/ảnh chụp màn hình — đa ngôn ngữ, giữ nguyên cấu trúc file gốc.")


def get_secret(key: str) -> str:
    """Đọc giá trị từ .streamlit/secrets.toml nếu có; trả về '' nếu không có file/khóa đó."""
    try:
        return st.secrets.get(key, "")
    except Exception:
        return ""


# ---- Cổng mật khẩu tùy chọn — chỉ bật khi APP_PASSWORD được cấu hình trong secrets.toml ----
_app_password = get_secret("APP_PASSWORD")
if _app_password:
    if "authed" not in st.session_state:
        st.session_state.authed = False
    if not st.session_state.authed:
        pwd = st.text_input("Nhập mật khẩu truy cập", type="password")
        if pwd == _app_password:
            st.session_state.authed = True
            st.rerun()
        elif pwd:
            st.error("Sai mật khẩu.")
        st.stop()

with st.sidebar:
    st.header("Cấu hình")
    _secret_key = get_secret("ANTHROPIC_API_KEY")
    if _secret_key:
        api_key = _secret_key
        st.success("🔑 API key đã được cấu hình sẵn.")
    else:
        api_key = st.text_input("Anthropic API key", type="password",
                                 help="Dạng sk-ant-... — lấy tại console.anthropic.com")
    st.markdown("---")

    lang_keys = list(LANGUAGES_VI.keys())
    src_options = ["auto"] + lang_keys + ["other"]
    tgt_options = lang_keys + ["other"]

    src = st.selectbox("Ngôn ngữ nguồn", src_options,
                        format_func=lambda k: "Tự động nhận diện" if k == "auto"
                        else ("Khác (tự nhập)" if k == "other" else LANGUAGES_VI[k]))
    src_custom = st.text_input("→ Nhập tên ngôn ngữ nguồn (tiếng Anh)", disabled=(src != "other"))

    tgt = st.selectbox("Ngôn ngữ đích", tgt_options,
                        format_func=lambda k: "Khác (tự nhập)" if k == "other" else LANGUAGES_VI[k])
    tgt_custom = st.text_input("→ Nhập tên ngôn ngữ đích (tiếng Anh)", disabled=(tgt != "other"))

    st.markdown("---")
    bilingual = st.checkbox("🈯 Dịch song ngữ (giữ cả bản gốc + bản dịch)", value=False)
    st.markdown("---")
    st.caption("Chuyên ngành: tối ưu cho kỹ sư MEP — giữ nguyên mã thiết bị, tiêu chuẩn "
               "(TCVN/QCVN/ASHRAE/NFPA) và đơn vị kỹ thuật.")


def resolve_lang_label(code, custom_text, table):
    if code == "other":
        return custom_text.strip() if custom_text.strip() else "the target language"
    return table[code]


uploaded = st.file_uploader(
    "Tải lên tệp cần dịch (PDF, DOCX, XLSX/XLS, JPG, PNG — kể cả ảnh chụp màn hình)",
    type=["pdf", "docx", "xlsx", "xls", "jpg", "jpeg", "png", "webp"],
)

if "result" not in st.session_state:
    st.session_state.result = None

page_from, page_to = None, None

if uploaded is not None:
    kind = uploaded.name.lower().rsplit(".", 1)[-1]
    kind = "xlsx" if kind in ("xlsx", "xls") else ("image" if kind in ("jpg", "jpeg", "png", "webp") else kind)
    file_bytes = uploaded.getvalue()

    st.write(f"**Tệp:** {uploaded.name} · **Loại:** {kind.upper()} · **Kích thước:** {len(file_bytes)/1024:.0f} KB")

    if kind == "pdf":
        try:
            n_pages = get_pdf_page_count(file_bytes)
            st.write(f"**Số trang:** {n_pages}")
            if n_pages > PDF_PAGE_WARN_THRESHOLD:
                st.warning(f"PDF có {n_pages} trang — dịch toàn bộ có thể mất khá lâu và tốn "
                           "nhiều lượt gọi API. Bạn có thể chọn khoảng trang cần dịch bên dưới.")
            c1, c2 = st.columns(2)
            page_from = c1.number_input("Từ trang", min_value=1, max_value=n_pages, value=1)
            page_to = c2.number_input("Đến trang", min_value=1, max_value=n_pages, value=n_pages)
        except Exception as e:
            st.error(f"Không đọc được PDF: {e}")

    if st.button("🌐 Dịch tài liệu", type="primary", disabled=not api_key):
        if not api_key:
            st.error("Vui lòng nhập Anthropic API key ở thanh bên trái.")
        else:
            try:
                client = get_client(api_key)
                src_label = ("ngon ngu nguon (tu dong nhan dien)" if src == "auto"
                              else resolve_lang_label(src, src_custom, LANGUAGES_EN))
                tgt_label = resolve_lang_label(tgt, tgt_custom, LANGUAGES_EN)
                system_prompt = build_system_prompt(src_label, tgt_label)

                progress = st.progress(0, text="Đang xử lý...")

                def cb(pct):
                    progress.progress(min(100, pct), text=f"Đang dịch... {pct}%")

                if kind == "pdf":
                    original, translated, out_bytes = translate_pdf_to_docx(
                        client, system_prompt, file_bytes, int(page_from), int(page_to), bilingual, cb)
                    st.session_state.result = {"kind": "pdf", "original": original, "translated": translated,
                                                "docx_bytes": out_bytes, "name": uploaded.name}

                elif kind == "docx":
                    original, translated, out_bytes = translate_docx(
                        client, system_prompt, file_bytes, bilingual, cb)
                    st.session_state.result = {"kind": "docx", "original": original, "translated": translated,
                                                "docx_bytes": out_bytes, "name": uploaded.name}

                elif kind == "xlsx":
                    original, translated, out_bytes = translate_xlsx(
                        client, system_prompt, file_bytes, bilingual, cb)
                    st.session_state.result = {"kind": "xlsx", "original": original, "translated": translated,
                                                "xlsx_bytes": out_bytes, "name": uploaded.name}

                elif kind == "image":
                    media_type = uploaded.type or "image/png"
                    translated = translate_image(client, system_prompt, tgt_label, file_bytes, media_type, bilingual)
                    cb(100)
                    st.session_state.result = {"kind": "image", "original": "(nội dung ảnh — xem tab Bản dịch)",
                                                "translated": translated, "name": uploaded.name}

                progress.progress(100, text="Hoàn tất")
                st.success("Đã dịch xong.")
            except Exception as e:
                st.error(f"Lỗi khi dịch: {e}")

if st.session_state.result:
    res = st.session_state.result
    tab1, tab2 = st.tabs(["📄 Văn bản gốc", "🌐 Bản dịch"])
    with tab1:
        st.text_area("Gốc", res["original"], height=420, label_visibility="collapsed")
    with tab2:
        st.text_area("Dịch", res["translated"], height=420, label_visibility="collapsed")

        base_name = res["name"].rsplit(".", 1)[0]
        st.download_button("⬇️ Tải bản dịch (.txt)", res["translated"], file_name=f"{base_name}_translated.txt")

        if res["kind"] in ("docx", "pdf") and "docx_bytes" in res:
            label = "⬇️ Tải Word đã dịch (.docx)" if res["kind"] == "docx" else "⬇️ Tải Word tái dựng từ PDF (.docx)"
            st.download_button(label, res["docx_bytes"], file_name=f"{base_name}_translated.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        if res["kind"] == "xlsx" and "xlsx_bytes" in res:
            st.download_button("⬇️ Tải bảng tính đã dịch (.xlsx)", res["xlsx_bytes"],
                                file_name=f"{base_name}_translated.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
