import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
from openpyxl import load_workbook

# =========================
# CONFIG UI
# =========================
st.set_page_config(page_title="OCR PDF Tool", layout="wide")

st.title("📄 OCR PDF → Excel (SM + Ngày)")

# =========================
# OCR FUNCTION
# =========================
def process_page(img):
    # chỉ OCR 1 lần (không xoay nữa → nhanh hơn)
    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )

    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    if sm and date:
        return sm.group(1), date.group(1)

    return None, None


# =========================
# EXTRACT PDF (OPTIMIZED)
# =========================
def extract_pdf(uploaded_file):
    results = []

    # giảm DPI → tăng tốc
    images = convert_from_bytes(uploaded_file.read(), dpi=150)

    progress = st.progress(0)
    status = st.empty()

    total = len(images)

    for i, img in enumerate(images, start=1):
        status.text(f"⚡ Đang xử lý {i}/{total}...")

        # 🔥 crop phần trên (tăng tốc cực mạnh)
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date
            })

        progress.progress(i / total)

    status.text("✅ Hoàn tất")
    return results


# =========================
# AUTO WIDTH EXCEL
# =========================
def auto_width(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_len + 3

    wb.save(excel_path)


# =========================
# MAIN UI
# =========================
uploaded_file = st.file_uploader("📤 Upload PDF", type=["pdf"])

if uploaded_file:
    if st.button("🚀 Xử lý"):
        with st.spinner("Đang OCR..."):
            data = extract_pdf(uploaded_file)

        if not data:
            st.error("❌ Không tìm thấy dữ liệu")
        else:
            df = pd.DataFrame(data)
            df.insert(0, "STT", range(1, len(df) + 1))

            # lưu file tạm
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                df.to_excel(tmp.name, index=False)
                auto_width(tmp.name)

                with open(tmp.name, "rb") as f:
                    st.download_button(
                        "📥 Tải file Excel",
                        f,
                        file_name="output.xlsx"
                    )

            st.success("🎉 Hoàn tất!")
