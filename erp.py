import io, os, re, math
import numpy as np
import pandas as pd
import streamlit as st
from typing import List, Dict, Tuple
import cv2, pytesseract
from PIL import Image

# --- Multiple-file upload version ---

st.set_page_config(page_title="ERP Translation Audit (Multi-View)", layout="wide")
st.title("ERP Translation Audit — Multiple Screenshot Support")

with st.expander("Optional: Tesseract path (Windows)"):
    tess_path = st.text_input("Full path to tesseract.exe", value="")
    if tess_path:
        pytesseract.pytesseract.tesseract_cmd = tess_path

col_inputs = st.columns(3)
with col_inputs[0]:
    xls_file = st.file_uploader("Upload Excel (sheet: 'Translations 2nd DRAFT', D=Greek, E=English)", type=["xlsx"])
with col_inputs[1]:
    imgs_gr = st.file_uploader("Upload Greek screenshots (one or more)", type=["png","jpg","jpeg"], accept_multiple_files=True)
with col_inputs[2]:
    imgs_en = st.file_uploader("Upload English screenshots (one or more)", type=["png","jpg","jpeg"], accept_multiple_files=True)


def ocr_boxes(file) -> pd.DataFrame:
    """OCR for one image, returns text/conf/x/y"""
    pil = Image.open(file).convert("RGB")
    img = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
    h, w = img.shape[:2]

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 7, 75, 75)
    gray = cv2.equalizeHist(gray)
    thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,45,10)

    data = pytesseract.image_to_data(
        thr,
        lang="eng+ell+spa",
        output_type=pytesseract.Output.DATAFRAME,
        config="--psm 6 --oem 3"
    )

    data = data.dropna(subset=["text"])
    data["text"] = data["text"].astype(str).str.strip()
    data = data[(data["text"]!="") & (data["conf"]!=-1)]
    if data.empty:
        return pd.DataFrame(columns=["text","conf","x","y"])
    data["x"] = (data["left"] + data["width"]/2) / w
    data["y"] = (data["top"] + data["height"]/2) / h
    return data[["text","conf","x","y"]].reset_index(drop=True)


def merge_ocr_results(files):
    """Combine multiple OCR outputs into one unified DataFrame"""
    frames = []
    for f in files:
        with st.spinner(f"OCR → {f.name}"):
            df = ocr_boxes(f)
            if not df.empty:
                df["source"] = f.name
                frames.append(df)
    if not frames:
        return pd.DataFrame(columns=["text","conf","x","y","source"])
    merged = pd.concat(frames, ignore_index=True)
    merged.drop_duplicates(subset=["text","x","y"], inplace=True)
    return merged


# Run button
run = st.button("Run audit", type="primary", use_container_width=True)

if run:
    if not (xls_file and imgs_gr and imgs_en):
        st.error("Please upload: Excel + Greek + English screenshots.")
        st.stop()

    with st.spinner("Loading Excel…"):
        trans_df = pd.read_excel(xls_file, sheet_name="Translations 2nd DRAFT", usecols="D:E")
        trans_df.columns = ["Greek","English"]

    # OCR all Greek images and merge
    with st.spinner("Processing Greek screenshots…"):
        gr_df = merge_ocr_results(imgs_gr)
        st.write(f"Greek tokens detected: {len(gr_df)}")

    # OCR all English images and merge
    with st.spinner("Processing English screenshots…"):
        en_df = merge_ocr_results(imgs_en)
        st.write(f"English tokens detected: {len(en_df)}")

    # now pass gr_df and en_df into your existing build_report() function
    with st.spinner("Matching & building report…"):
        report = build_report(trans_df, gr_df, en_df)

    st.success("✅ Audit complete")
    st.dataframe(report, use_container_width=True, height=420)

    # Download button
    buf = io.BytesIO()
    report.to_excel(buf, index=False)
    buf.seek(0)
    st.download_button(
        "Download Excel Report",
        data=buf,
        file_name="erp_translation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
