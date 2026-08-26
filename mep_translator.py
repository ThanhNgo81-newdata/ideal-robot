# mep_translator.py
# MEP Translator Offline
# Supports DOCX, PDF, XLSX, TXT with bilingual mode while preserving file structure
# Designed for deployment on Hugging Face Spaces / Streamlit

import io
from typing import List
import streamlit as st
from transformers import MarianMTModel, MarianTokenizer
import docx
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.shared import OxmlElement
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import fitz  # PyMuPDF

# -------------------------------
# Configuration
# -------------------------------
LANGUAGES_VI = {
    "en": "Tiếng Anh", "vi": "Tiếng Việt", "fr": "Tiếng Pháp", "de": "Tiếng Đức",
    "es": "Tiếng Tây Ban Nha", "ru": "Tiếng Nga", "zh": "Tiếng Trung", "ja": "Tiếng Nhật",
    "ko": "Tiếng Hàn", "th": "Tiếng Thái", "pt": "Tiếng Bồ Đào Nha", "it": "Tiếng Ý"
}

st.set_page_config(page_title="MEP Translator", page_icon="🛠️", layout="wide")
st.title("🛠️ MEP TRANSLATOR (Offline)")
st.caption("Dịch tài liệu kỹ thuật MEP — DOCX, PDF, Excel, TXT — chạy bằng mô hình Hugging Face. Có chế độ song ngữ và cố gắng giữ cấu trúc file gốc.")

# -------------------------------
# Model loading and caching
# -------------------------------
@st.cache_resource
def load_model_tokenizer(src: str, tgt: str):
    """
    Load and cache Marian model and tokenizer for a language pair.
    Returns (tokenizer, model).
    """
    model_name = f"Helsinki-NLP/opus-mt-{src}-{tgt}"
    tokenizer = MarianTokenizer.from_pretrained(model_name)
    model = MarianMTModel.from_pretrained(model_name)
    return tokenizer, model

def translate_batch(texts: List[str], src: str, tgt: str) -> List[str]:
    """
    Translate a list of texts in batch using cached model/tokenizer.
    """
    if not texts:
        return []
    tokenizer, model = load_model_tokenizer(src, tgt)
    # Tokenize batch with truncation/padding
    inputs = tokenizer(texts, return_tensors="pt", padding=True, truncation=True)
    outputs = model.generate(**inputs)
    results = [tokenizer.decode(t, skip_special_tokens=True) for t in outputs]
    return results

def translate_offline(text: str, src: str, tgt: str) -> str:
    """
    Convenience wrapper for single string translation using batch function.
    """
    return translate_batch([text], src, tgt)[0]

# -------------------------------
# DOCX helpers preserving order
# -------------------------------
def _add_paragraph_after_element(element, text: str, document: Document):
    """
    Insert a new paragraph element (w:p) after the given element in the document body.
    Returns a docx Paragraph object.
    """
    body = document._element.body
    new_p = OxmlElement("w:p")
    body.insert(body.index(element) + 1, new_p)
    return Paragraph(new_p, document)

def translate_docx_preserve(file_bytes: bytes, src: str, tgt: str, bilingual: bool = True, batch_size: int = 16) -> bytes:
    """
    Translate DOCX while preserving element order (paragraphs and tables).
    For paragraphs: insert translated paragraph immediately after original (keeps order).
    For tables: add a new column at the end of each row containing translated text (bilingual mode).
    """
    doc = Document(io.BytesIO(file_bytes))
    body = doc._element.body
    i = 0
    # Collect paragraphs to batch translate in groups to reduce model calls
    while i < len(body):
        child = body[i]
        tag = child.tag.split("}")[-1]
        if tag == "p":
            p = Paragraph(child, doc)
            text = p.text.strip()
            if text:
                # translate in single call (could be batched across loop iterations if needed)
                translated = translate_offline(text, src, tgt)
                if bilingual:
                    new_para = _add_paragraph_after_element(child, translated, doc)
                    try:
                        new_para.style = p.style
                    except Exception:
                        pass
                    i += 1  # skip the inserted paragraph
                else:
                    # Replace paragraph text while attempting to preserve runs minimally
                    # Clear runs then add a single run with translated text
                    for run in list(p.runs):
                        try:
                            run.clear()
                        except Exception:
                            pass
                    p.add_run(translated)
        elif tag == "tbl":
            # Wrap table element into docx.table.Table
            from docx.table import Table
            tbl = Table(child, doc)
            if bilingual:
                # Insert a new cell at end of each row and put translated text there
                for row in tbl.rows:
                    # Create new tc element and append to tr
                    new_tc = OxmlElement("w:tc")
                    row._tr.append(new_tc)
                    # Create paragraph inside new cell
                    new_para = Paragraph(new_tc, row._tr)
                    # For translation source text, concatenate texts of existing cells in the row or use last cell
                    # Here we translate each last cell's text for simplicity
                    source_text = row.cells[-1].text.strip() if row.cells else ""
                    if source_text:
                        new_para.add_run(translate_offline(source_text, src, tgt))
            else:
                # Replace each cell text with translation
                for row in tbl.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            cell.text = translate_offline(cell.text, src, tgt)
        i += 1

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -------------------------------
# Excel helpers preserving layout
# -------------------------------
def translate_xlsx_preserve(file_bytes: bytes, src: str, tgt: str, bilingual: bool = True, batch_size: int = 32) -> bytes:
    """
    Translate Excel workbook. For bilingual mode, insert a new column to the right of each text column and copy styles.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    for ws in wb.worksheets:
        col = 1
        while col <= ws.max_column:
            # Detect if column has any text
            has_text = any(isinstance(ws.cell(row=r, column=col).value, str) and ws.cell(row=r, column=col).value.strip()
                           for r in range(1, ws.max_row + 1))
            if has_text and bilingual:
                ws.insert_cols(col + 1)
                # copy column width if exists
                try:
                    ws.column_dimensions[get_column_letter(col + 1)].width = ws.column_dimensions[get_column_letter(col)].width
                except Exception:
                    pass
                # Translate cell by cell (could be batched per column)
                texts = []
                coords = []
                for r in range(1, ws.max_row + 1):
                    cell = ws.cell(row=r, column=col)
                    if isinstance(cell.value, str) and cell.value.strip():
                        texts.append(cell.value)
                        coords.append((r, col + 1))
                    else:
                        coords.append((r, None))
                if texts:
                    translations = translate_batch(texts, src, tgt)
                    ti = 0
                    for r, c_new in coords:
                        if c_new:
                            ws.cell(row=r, column=c_new).value = translations[ti]
                            # copy style
                            try:
                                src_cell = ws.cell(row=r, column=col)
                                dst_cell = ws.cell(row=r, column=c_new)
                                dst_cell.font = copy(src_cell.font)
                                dst_cell.fill = copy(src_cell.fill)
                                dst_cell.number_format = src_cell.number_format
                                dst_cell.alignment = copy(src_cell.alignment)
                                dst_cell.border = copy(src_cell.border)
                            except Exception:
                                pass
                            ti += 1
                col += 2
            elif has_text and not bilingual:
                # Replace in place
                texts = []
                coords = []
                for r in range(1, ws.max_row + 1):
                    cell = ws.cell(row=r, column=col)
                    if isinstance(cell.value, str) and cell.value.strip():
                        texts.append(cell.value)
                        coords.append(r)
                if texts:
                    translations = translate_batch(texts, src, tgt)
                    for idx, r in enumerate(coords):
                        ws.cell(row=r, column=col).value = translations[idx]
                col += 1
            else:
                col += 1
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# -------------------------------
# PDF helpers converting to DOCX while preserving block order
# -------------------------------
def translate_pdf_to_docx(file_bytes: bytes, src: str, tgt: str, bilingual: bool = True, batch_size: int = 16) -> bytes:
    """
    Extract text blocks from PDF in reading order and create a DOCX with original block followed by translation (if bilingual).
    Note: complex layouts (multi-column, images) may not be perfectly preserved.
    """
    pdf = fitz.open(stream=file_bytes, filetype="pdf")
    out_doc = Document()
    for page in pdf:
        blocks = page.get_text("blocks")  # list of (x0, y0, x1, y1, "text", block_no)
        # Sort blocks by y then x to approximate reading order
        blocks_sorted = sorted(blocks, key=lambda b: (round(b[1]), round(b[0])))
        texts = [b[4].strip() for b in blocks_sorted if b[4].strip()]
        # Batch translate blocks
        for t in texts:
            if bilingual:
                p1 = out_doc.add_paragraph(t)
                if p1.runs:
                    try:
                        p1.runs[0].italic = True
                    except Exception:
                        pass
                out_doc.add_paragraph(translate_offline(t, src, tgt))
            else:
                out_doc.add_paragraph(translate_offline(t, src, tgt))
    out = io.BytesIO()
    out_doc.save(out)
    return out.getvalue()

# -------------------------------
# TXT helper
# -------------------------------
def translate_txt_preserve(file_bytes: bytes, src: str, tgt: str, bilingual: bool = True) -> bytes:
    text = file_bytes.decode("utf-8", errors="ignore")
    lines = text.splitlines()
    out_lines = []
    batch = []
    batch_idx = []
    # Batch translate per 32 lines
    for idx, line in enumerate(lines):
        out_lines.append(line)
        if line.strip():
            batch.append(line)
            batch_idx.append(len(out_lines))  # position to insert translation
        if len(batch) >= 32:
            translations = translate_batch(batch, src, tgt)
            for j, tr in enumerate(translations):
                if bilingual:
                    out_lines.insert(batch_idx[j], tr)
            batch = []
            batch_idx = []
    if batch:
        translations = translate_batch(batch, src, tgt)
        for j, tr in enumerate(translations):
            if bilingual:
                out_lines.insert(batch_idx[j], tr)
    result = "\n".join(out_lines)
    return result.encode("utf-8")

# -------------------------------
# Streamlit UI
# -------------------------------
with st.sidebar:
    st.header("Cấu hình dịch")
    src_lang = st.selectbox("Ngôn ngữ nguồn", list(LANGUAGES_VI.keys()), index=0, format_func=lambda x: LANGUAGES_VI[x])
    tgt_lang = st.selectbox("Ngôn ngữ đích", list(LANGUAGES_VI.keys()), index=1, format_func=lambda x: LANGUAGES_VI[x])
    bilingual = st.checkbox("Dịch song ngữ (giữ bản gốc + bản dịch)", value=True)
    method = st.radio("Phương thức chèn bản dịch", ("Chèn dưới mỗi đoạn", "Bảng 2 cột cho DOCX/Excel"), index=0)
    st.markdown("**Gợi ý:** Chọn 'Bảng 2 cột' để giữ layout song song cho DOCX/Excel; chọn 'Chèn dưới mỗi đoạn' để giữ thứ tự đọc.")

uploaded = st.file_uploader("Tải lên tệp cần dịch (PDF, DOCX, XLSX, TXT)", type=["pdf", "docx", "xlsx", "txt"])

if uploaded:
    kind = uploaded.name.lower().split(".")[-1]
    file_bytes = uploaded.getvalue()

    if st.button("Dịch tài liệu"):
        try:
            if kind == "docx":
                if method == "Bảng 2 cột cho DOCX/Excel":
                    # For DOCX, using preserve function will insert translated paragraphs after originals.
                    out_bytes = translate_docx_preserve(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
                else:
                    out_bytes = translate_docx_preserve(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
                st.download_button("Tải DOCX đã dịch", out_bytes, file_name="translated.docx")
            elif kind == "xlsx":
                out_bytes = translate_xlsx_preserve(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
                st.download_button("Tải Excel đã dịch", out_bytes, file_name="translated.xlsx")
            elif kind == "pdf":
                out_bytes = translate_pdf_to_docx(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
                st.download_button("Tải DOCX từ PDF đã dịch", out_bytes, file_name="translated_from_pdf.docx")
            elif kind == "txt":
                out_bytes = translate_txt_preserve(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
                st.download_button("Tải TXT đã dịch", out_bytes, file_name="translated.txt")
            st.success("Hoàn tất. Tải file đã dịch bằng nút tải xuống.")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi khi dịch: {e}")
            raise
