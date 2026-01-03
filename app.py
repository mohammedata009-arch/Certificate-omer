import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import zipfile
from PIL import Image

st.set_page_config(page_title="نظام شهادات TTC", layout="wide")

PASSWORD = "TTC_2024"

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    pwd = st.sidebar.text_input("كلمة المرور:", type="password")
    if st.sidebar.button("دخول"):
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("خطأ!")
    st.stop()

st.title("🎓 لوحة تحكم شهادات مركز مدربون")
col_in, col_pre = st.columns([1, 1])

with col_in:
    excel_file = st.file_uploader("1. ارفع ملف الإكسل", type=['xlsx'])
    sig_file = st.file_uploader("2. ارفع التوقيع (PNG)", type=['png'])
    font_size = st.slider("حجم خط الاسم", 15, 35, 22)

if excel_file:
    df = pd.read_excel(excel_file)
    with col_pre:
        test_row = df.iloc[0]
        doc = fitz.open("OMER.pdf")
        page = doc[0]
        page.insert_text((128.45, 430.0), str(test_row['Name']).upper(), fontsize=font_size, fontname="helv-bold")
        page.insert_text((245.30, 535.40), str(test_row['Course']), fontsize=18, fontname="helv-bold")
        page.insert_text((475.0, 792.0), str(test_row['CertNo']), fontsize=11, fontname="helv")
        if sig_file:
            page.insert_image(fitz.Rect(430, 700, 550, 750), stream=sig_file.getvalue())
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        st.image(Image.frombytes("RGB", [pix.width, pix.height], pix.samples), use_container_width=True)

    if st.button("🚀 إصدار الشهادات وتحميلها (ZIP)"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for idx, row in df.iterrows():
                doc = fitz.open("OMER.pdf")
                p = doc[0]
                p.insert_text((128.45, 430.0), str(row['Name']).upper(), fontsize=font_size, fontname="helv-bold")
                p.insert_text((245.30, 535.40), str(row['Course']), fontsize=18, fontname="helv-bold")
                p.insert_text((475.0, 792.0), str(row['CertNo']), fontsize=11, fontname="helv")
                if sig_file:
                    p.insert_image(fitz.Rect(430, 700, 550, 750), stream=sig_file.getvalue())
                zip_file.writestr(f"{row['Name']}.pdf", doc.write())
                doc.close()
        st.download_button("⬇️ تحميل الكل", data=zip_buffer.getvalue(), file_name="Certs.zip")
