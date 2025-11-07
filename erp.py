import io, os, re, math
import numpy as np
import pandas as pd
import streamlit as st
import cv2, pytesseract
from PIL import Image
from rapidfuzz import fuzz

# ==============================
# CONFIG
# ==============================
st.set_page_config(page_title="ERP Translation Audit", layout="wide")
st.title("🔍 ERP Translation Audit — Translation Detective")

with st.expander("Optional: Tesseract path (Windows)"):
    tess_path = st.text_input("Full path to tesseract.exe", value="")
    if tess_path:
        pytesseract.pytesseract.tesseract_cmd = tess_path

# ==============================
# UPLOADS
# ==============================
col_inputs = st.columns(3)
with col_inputs[0]:
    xls_file = st.file_uploader(
        "Upload Excel (sheet: 'Translations 2nd DRAFT', D=Greek, E=English)",
        type=["xlsx"]
    )
with col_inputs[1]:
    imgs_gr = st.file_uploader(
        "Upload Greek screenshots (one or more)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )
with col_inputs[2]:
    imgs_en = st.file_uploader(
        "Upload English screenshots (one or more)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )

# ==============================
# OCR FUNCTIONS
# ==============================
def ocr_boxes(file) -> pd.DataFrame:
    """Run OCR on one screenshot and return text/conf/x/y/filename."""
    pil = Image.open(file).convert("RGB")
    img = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
    h, w = img.shape[:2]

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 7, 75, 75)
    gray = cv2.equalizeHist(gray)
    thr = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 45, 10
    )

    data = pytesseract.image_to_data(
        thr,
        lang="eng+ell+spa",
        output_type=pytesseract.Output.DATAFRAME,
        config="--psm 6 --oem 3"
    )

    data = data.dropna(subset=["text"])
    data["text"] = data["text"].astype(str).str.strip()
    data = data[(data["text"] != "") & (data["conf"] != -1)]
    if data.empty:
        return pd.DataFrame(columns=["text", "conf", "x", "y", "source"])

    data["x"] = (data["left"] + data["width"] / 2) / w
    data["y"] = (data["top"] + data["height"] / 2) / h
    data["source"] = file.name
    return data[["text", "conf", "x", "y", "source"]].reset_index(drop=True)


def merge_ocr_results(files):
    """Combine OCR outputs from multiple screenshots into one dataset."""
    frames = []
    for f in files:
        with st.spinner(f"OCR → {f.name}"):
            df = ocr_boxes(f)
            if not df.empty:
                frames.append(df)
    if not frames:
        return pd.DataFrame(columns=["text", "conf", "x", "y", "source"])
    merged = pd.concat(frames, ignore_index=True)
    merged.drop_duplicates(subset=["text", "x", "y"], inplace=True)
    return merged


# ==============================
# TRANSLATION DETECTIVE CORE
# ==============================
def build_report(trans_df: pd.DataFrame, gr_df: pd.DataFrame, en_df: pd.DataFrame) -> pd.DataFrame:
    """Compare Greek vs English OCR tokens and produce a clean audit table."""
    rows = []
    for _, tr in trans_df.iterrows():
        greek_expected = str(tr["Greek"]).strip()
        eng_expected = str(tr["English"]).strip()

        # --- Protect regex characters ---
        greek_pattern = re.escape(greek_expected[:4])
        english_pattern = re.escape(eng_expected[:4])

        # --- Find OCR detections safely ---
        gr_hits = gr_df[gr_df["text"].str.contains(greek_pattern, case=False, na=False, regex=True)]
        en_hits = en_df[en_df["text"].str.contains(english_pattern, case=False, na=False, regex=True)]

        if gr_hits.empty:
            rows.append({
                "Greek (Expected)": greek_expected,
                "Greek (Detected)": "",
                "English (Expected)": eng_expected,
                "English (Detected)": "",
                "Status": 4,
                "Status Description": "Field Not Found on the Report View",
                "Confidence": 0.0,
                "Reason": "Greek header not visible in screenshots",
                "Screenshot": ""
            })
            continue

        # take first Greek detection
        g = gr_hits.iloc[0]
        gtxt, gx, gy, gconf, gsrc = g["text"], g["x"], g["y"], g["conf"], g.get("source", "")

        # find best English match
        best_row, best_score = None, 0
        for _, e in en_df.iterrows():
            sim_txt = fuzz.token_set_ratio(gtxt, e["text"]) / 100
            dist = math.sqrt((gx - e["x"]) ** 2 + (gy - e["y"]) ** 2)
            pos_score = max(0, 1 - dist)
            total = 0.7 * sim_txt + 0.3 * pos_score
            if total > best_score:
                best_score = total
                best_row = e

        if best_row is None or best_score < 0.65:
            rows.append({
                "Greek (Expected)": greek_expected,
                "Greek (Detected)": gtxt,
                "English (Expected)": eng_expected,
                "English (Detected)": "",
                "Status": 3,
                "Status Description": "Field Not Translated",
                "Confidence": round(best_score, 3),
                "Reason": "No strong English match detected",
                "Screenshot": gsrc
            })
            continue

        e = best_row
        if best_score >= 0.9:
            estatus, reason = 1, "Good match"
        elif best_score >= 0.75:
            estatus, reason = 2, "Possible synonym / slightly different"
        else:
            estatus, reason = 0, "Low certainty match"

        rows.append({
            "Greek (Expected)": greek_expected,
            "Greek (Detected)": gtxt,
            "English (Expected)": eng_expected,
            "English (Detected)": e["text"],
            "Status": estatus,
            "Status Description": {
                1: "Translated_Correct",
                2: "Translated_Not Accurate",
                0: "Pending Review",
                3: "Field Not Translated",
                4: "Field Not Found on the Report View"
            }[estatus],
            "Confidence": round(best_score, 3),
            "Reason": reason,
            "Screenshot": e.get("source", "")
        })

    out = pd.DataFrame(rows)
    order = pd.CategoricalDtype([
        "Field Not Found on the Report View",
        "Field Not Translated",
        "Translated_Not Accurate",
        "Pending Review",
        "Translated_Correct"
    ], ordered=True)
    out["Status Description"] = out["Status Description"].astype(order)
    out = out.sort_values(["Status Description", "Greek (Expected)"]).reset_index(drop=True)
    return out


# ==============================
# RUN PIPELINE
# ==============================
run = st.button("🚀 Run Audit", type="primary", use_container_width=True)

if run:
    if not (xls_file and imgs_gr and imgs_en):
        st.error("Please upload: Excel + Greek + English screenshots.")
        st.stop()

    with st.spinner("Loading Excel…"):
        trans_df = pd.read_excel(xls_file, sheet_name="Translations 2nd DRAFT", usecols="D:E")
        trans_df.columns = ["Greek", "English"]

    with st.spinner("Processing Greek screenshots…"):
        gr_df = merge_ocr_results(imgs_gr)
        st.write(f"✅ Greek tokens detected: {len(gr_df)}")

    with st.spinner("Processing English screenshots…"):
        en_df = merge_ocr_results(imgs_en)
        st.write(f"✅ English tokens detected: {len(en_df)}")

    if gr_df.empty or en_df.empty:
        st.error("⚠️ OCR did not detect any text. Please upload clear screenshots.")
        st.stop()

    with st.spinner("Analyzing translations…"):
        report = build_report(trans_df, gr_df, en_df)

    st.success("✅ Audit complete")
    st.dataframe(report, use_container_width=True, height=420)

    buf = io.BytesIO()
    report.to_excel(buf, index=False)
    buf.seek(0)
    st.download_button(
        "📥 Download Excel Report",
        data=buf,
        file_name="erp_translation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
