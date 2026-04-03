import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import zipfile
import os
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="OCR Drive UI", layout="wide")

# =========================
# SESSION
# =========================
if "processing" not in st.session_state:
    st.session_state.processing = False

if "done" not in st.session_state:
    st.session_state.done = False

if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False

# =========================
# STYLE (CLEAN UI)
# =========================
st.markdown("""
<style>
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.stApp { background: #f8fafc; }

.block-container {
    padding-top: 0.5rem !important;
}

/* HEADER */
.header {
    padding:10px 0;
    font-size:20px;
    font-weight:600;
}

/* UPLOADER */
[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5f5;
    padding: 30px;
    border-radius: 16px;
    text-align: center;
    background: white;
}

/* FILE ROW */
.file-row {
    font-size:14px;
    margin-top:10px;
}

.file-name {
    font-weight:500;
}

.file-status {
    color:#64748b;
    font-size:13px;
}

/* PROGRESS */
.progress {
    height:6px;
    background:#e5e7eb;
    border-radius:10px;
    overflow:hidden;
    margin-top:4px;
}

.progress-bar {
    height:100%;
    background:linear-gradient(90deg,#0ea5e9,#22c55e);
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">📁 CHECK PDF TO EXCEL ( SM ) </div>', unsafe_allow_html=True)

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

# =========================
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# PROCESS
# =========================
def extract_pdf(file, box, idx, total, global_bar):
    results = []
    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):
        percent = int((i/total_pages)*100)
        global_percent = int(((idx + i/total_pages)/total)*100)

        html = f"""
<div class="file-row">
    <div class="file-name">📄 {file.name}</div>
    <div class="file-status">Trang {i}/{total_pages} • {percent}%</div>
    <div class="progress">
        <div class="progress-bar" style="width:{percent}%"></div>
    </div>
</div>
"""
        box.markdown(html, unsafe_allow_html=True)
        global_bar.progress(global_percent)

        # crop top
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)
        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    global_bar = st.progress(0)

    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Process Files"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            for i, f in enumerate(uploaded_files):

                data = extract_pdf(f, boxes[i], i, len(uploaded_files), global_bar)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    name = os.path.splitext(f.name)[0] + ".xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name, index=False)

                        wb = load_workbook(tmp.name)
                        ws = wb.active

                        for col in ws.columns:
                            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                            ws.column_dimensions[col[0].column_letter].width = max_len + 3

                        wb.save(tmp.name)
                        zipf.write(tmp.name, name)

        st.session_state.zip = zip_buffer.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 Xử lý xong!")

    with open(st.session_state.zip, "rb") as f:
        zip_data = f.read()

    if st.download_button(
        "📥 Download ZIP",
        zip_data,
        file_name="THLTOOL.zip",
        mime="application/zip"
    ):
        st.toast("✅ Download xong!", icon="🎉")

        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
