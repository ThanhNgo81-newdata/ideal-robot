# mep_translator.py
# Ứng dụng dịch tài liệu kỹ thuật MEP (DOCX, PDF, Excel, TXT)
# Chạy offline bằng Hugging Face, hỗ trợ chế độ song ngữ

import io
import streamlit as st
from transformers import MarianMTModel, MarianTokenizer
import docx
from docx.text.paragraph import Paragraph
import openpyxl
import fitz  # PyMuPDF

# ----------------------------------------------------------------------------
# Cấu hình chung
# ----------------------------------------------------------------------------

LANGUAGES_VI = {
    "en": "Tiếng Anh", "vi": "Tiếng Việt", "fr": "Tiếng Pháp", "de": "Tiếng Đức",
    "es": "Tiếng Tây Ban Nha", "ru": "Tiếng Nga", "zh": "Tiếng Trung", "ja": "Tiếng Nhật",
    "ko": "Tiếng Hàn", "th": "Tiếng Thái", "pt": "Tiếng Bồ Đào Nha", "it": "Tiếng Ý"
}

st.set_page_config(page_title="MEP Translator", page_icon="🛠️", layout="wide")
st.title("🛠️ MEP TRANSLATOR (Offline)")
st.caption("Dịch tài liệu kỹ thuật MEP — DOCX, PDF, Excel, TXT — chạy offline bằng Hugging Face, có chế độ song ngữ.")

# ----------------------------------------------------------------------------
# Hàm dịch offline Hugging Face
# ----------------------------------------------------------------------------
   
def translate_offline(text: str, src: str, tgt: str) -> str:
    model_name = f"Helsinki-NLP/opus-mt-{src}-{tgt}"
    tokenizer = MarianTokenizer.from_pretrained(model_name)
    model = MarianMTModel.from_pretrained(model_name)
    inputs = tokenizer(text, return_tensors="pt", padding=True)
    translated = model.generate(**inputs)
    return tokenizer.decode(translated[0], skip_special_tokens=True) 
    
# ----------------------------------------------------------------------------
# WORD (.docx) — dịch tại chỗ, giữ nguyên style/heading/bảng/bullet
# ----------------------------------------------------------------------------

def translate_docx(file_bytes, src, tgt, bilingual=True):
    d = docx.Document(io.BytesIO(file_bytes))
    for p in d.paragraphs:
        if p.text.strip():
            translated = translate_offline(p.text, src, tgt)
            if bilingual:
                new_p = insert_paragraph_after(p, translated)
                new_p.style = p.style
            else:
                p.text = translated
    out = io.BytesIO()
    d.save(out)
    return out.getvalue()

# ----------------------------------------------------------------------------
# EXCEL (.xlsx) — dịch tại chỗ, giữ nguyên sheet/style/merge/formula
# ----------------------------------------------------------------------------

def translate_xlsx(file_bytes, src, tgt, bilingual=True):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    for ws in wb.worksheets:
        max_col = ws.max_column
        for col_idx in range(max_col, 0, -1):
            has_text = any(isinstance(ws.cell(row=r, column=col_idx).value, str) 
                           and ws.cell(row=r, column=col_idx).value.strip()
                           for r in range(1, ws.max_row+1))
            if has_text and bilingual:
                ws.insert_cols(col_idx+1)
                for r in range(1, ws.max_row+1):
                    cell = ws.cell(row=r, column=col_idx)
                    if isinstance(cell.value, str) and cell.value.strip():
                        ws.cell(row=r, column=col_idx+1).value = translate_offline(cell.value, src, tgt)
            elif has_text and not bilingual:
                for r in range(1, ws.max_row+1):
                    cell = ws.cell(row=r, column=col_idx)
                    if isinstance(cell.value, str) and cell.value.strip():
                        cell.value = translate_offline(cell.value, src, tgt)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# -------------------------------
# PDF -> DOCX
# -------------------------------

def translate_pdf(file_bytes, src, tgt, bilingual=True):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    out_doc = docx.Document()
    for page in doc:
        blocks = page.get_text("blocks")
        texts = [b[4].strip() for b in blocks if b[4].strip()]
        for t in texts:
            translated = translate_offline(t, src, tgt)
            if bilingual:
                p1 = out_doc.add_paragraph(t)
                if p1.runs:
                    p1.runs[0].italic = True
                out_doc.add_paragraph(translated)
            else:
                out_doc.add_paragraph(translated)
    out = io.BytesIO()
    out_doc.save(out)
    return out.getvalue()

# ----------------------------------------------------------------------------
# Giao diện Streamlit
# ----------------------------------------------------------------------------

with st.sidebar:
    st.header("Cấu hình dịch")
    src_lang = st.selectbox("Ngôn ngữ nguồn", list(LANGUAGES_VI.keys()), format_func=lambda x: LANGUAGES_VI[x])
    tgt_lang = st.selectbox("Ngôn ngữ đích", list(LANGUAGES_VI.keys()), format_func=lambda x: LANGUAGES_VI[x])
    bilingual = st.checkbox("🈯 Dịch song ngữ (giữ cả bản gốc + bản dịch)", value=False)

uploaded = st.file_uploader("Tải lên tệp cần dịch (PDF, DOCX, XLSX, TXT)", type=["pdf","docx","xlsx","txt"])

if uploaded:
    kind = uploaded.name.lower().split(".")[-1]
    file_bytes = uploaded.getvalue()

    if st.button("🌐 Dịch tài liệu"):
        if kind == "docx":
            out_bytes = translate_docx(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
            st.download_button("⬇️ Tải DOCX đã dịch", out_bytes, file_name="translated.docx")
        elif kind == "xlsx":
            out_bytes = translate_xlsx(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
            st.download_button("⬇️ Tải Excel đã dịch", out_bytes, file_name="translated.xlsx")
        elif kind == "pdf":
            out_bytes = translate_pdf(file_bytes, src_lang, tgt_lang, bilingual=bilingual)
            st.download_button("⬇️ Tải DOCX từ PDF đã dịch", out_bytes, file_name="translated_from_pdf.docx")
        elif kind == "txt":
            text = file_bytes.decode("utf-8")
            translated = translate_offline(text, src_lang, tgt_lang)
            if bilingual:
                combined = f"{text}\n\n---\n\n{translated}"
                st.text_area("Bản dịch song ngữ", combined, height=400)
                st.download_button("⬇️ Tải TXT song ngữ", combined, file_name="translated_bilingual.txt")
            else:
                st.text_area("Bản dịch", translated, height=400)
                st.download_button("⬇️ Tải TXT đã dịch", translated, file_name="translated.txt")

