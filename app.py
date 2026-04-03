import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
from openpyxl import load_workbook

st.set_page_config(page_title="OCR PDF Tool", layout="wide")

st.title("📄 OCR PDF → Excel (SM + Ngày)")

def process_page(img):
    for angle in [0, 180]:
        rotated = img.rotate(angle, expand=True)

        text = pytesseract.image_to_string(
            rotated,
            lang='eng',
            config='--oem 3 --psm 6'
        )

        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

        if sm and date:
            return sm.group(1), date.group(1)

    return None, None


def extract_pdf(uploaded_file):
    results = []

    images = convert_from_bytes(uploaded_file.read(), dpi=200)

    for i, img in enumerate(images, start=1):
        st.write(f"👉 Trang {i}")

        sm, date = process_page(img)

        if sm and date:
            st.success(f"{sm} - {date}")
            results.append({
                "SM": sm,
                "Ngày": date
            })
        else:
            st.error("Không đọc được")

    return results


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


uploaded_file = st.file_uploader("📤 Upload PDF", type=["pdf"])

if uploaded_file:
    if st.button("🚀 Xử lý"):
        with st.spinner("Đang OCR..."):
            data = extract_pdf(uploaded_file)

        if not data:
            st.error("❌ Không có dữ liệu")
        else:
            df = pd.DataFrame(data)
            df.insert(0, "STT", range(1, len(df) + 1))

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                df.to_excel(tmp.name, index=False)
                auto_width(tmp.name)

                with open(tmp.name, "rb") as f:
                    st.download_button(
                        "📥 Tải Excel",
                        f,
                        file_name="output.xlsx"
                    )

            st.success("✅ Hoàn tất!")