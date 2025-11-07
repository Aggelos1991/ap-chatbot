# app.py — ERP Translation Audit (Streamlit)
import io, os, re, math
import numpy as np
import pandas as pd
import streamlit as st
from typing import List, Dict, Tuple

# ---------- OCR ----------
import cv2
import pytesseract
from PIL import Image

# ---------- Matching (primary: Sentence-Transformers, fallback: RapidFuzz) ----------
USE_ST = False
try:
    from sentence_transformers import SentenceTransformer
    from sklearn.metrics.pairwise import cosine_similarity
    ST_MODEL = SentenceTransformer("paraphrase-multilingual-mpnet-base-v2")
    USE_ST = True
except Exception:
    from rapidfuzz import fuzz

# ---------- App UI ----------
st.set_page_config(page_title="ERP Translation Audit", layout="wide")
st.title("ERP Translation Audit — Greek ↔ English (Headers)")

with st.expander("Optional: Tesseract path (Windows)"):
    tess_path = st.text_input("Full path to tesseract.exe (leave empty if on PATH)", value="")
    if tess_path:
        pytesseract.pytesseract.tesseract_cmd = tess_path

col_inputs = st.columns(3)
with col_inputs[0]:
    xls_file = st.file_uploader("Upload Excel (sheet: 'Translations 2nd DRAFT', D=Greek, E=English)", type=["xlsx"])
with col_inputs[1]:
    img_gr = st.file_uploader("Upload Greek UI screenshot (PNG/JPG)", type=["png","jpg","jpeg"])
with col_inputs[2]:
    img_en = st.file_uploader("Upload English UI screenshot (PNG/JPG)", type=["png","jpg","jpeg"])

# ---------- Status map ----------
STATUS_DESCRIPTIONS = {
    1: "Translated_Correct",
    2: "Translated_Not Accurate",
    3: "Field Not Translated",
    4: "Field Not Found on the Report View",
    0: "Pending Review",
    5: "Need review, Data on the field does not match the title"
}

THRESH_SEM_HIGH = 0.90
THRESH_SEM_MED  = 0.75
THRESH_OCR_CONF = 60.0  # tesseract conf is 0–100
WEIGHT_SEM = 0.7
WEIGHT_POS = 0.3

# ---------- Helpers ----------
def load_excel_pairs(file) -> pd.DataFrame:
    # Read only columns D:E (0-based is 3:4). Rename safely.
    df = pd.read_excel(file, sheet_name="Translations 2nd DRAFT", usecols="D:E", dtype=str)
    cols = list(df.columns)
    if len(cols) < 2:
        raise ValueError("Could not read columns D:E from the specified sheet.")
    df.columns = ["Greek", "English"]
    df["Greek"] = df["Greek"].fillna("").astype(str).str.strip()
    df["English"] = df["English"].fillna("").astype(str).str.strip()
    df = df[(df["Greek"]!="") | (df["English"]!="")]
    return df

def preprocess_for_ocr(img: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 5, 55, 55)
    thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,41,11)
    return thr

def ocr_boxes(file) -> pd.DataFrame:
    # Return tokens with text, conf, normalized x,y (0..1)
    pil = Image.open(file).convert("RGB")
    img = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
    h, w = img.shape[:2]
    proc = preprocess_for_ocr(img)

    data = pytesseract.image_to_data(proc, lang="eng+ell", output_type=pytesseract.Output.DATAFRAME)
    # Filter valid words
    data = data.dropna(subset=["text"])
    data["text"] = data["text"].astype(str).str.strip()
    data = data[(data["text"]!="") & (data["conf"]!=-1)]
    if data.empty:
        return pd.DataFrame(columns=["text","conf","x","y"])

    data["x"] = (data["left"] + data["width"]/2) / w
    data["y"] = (data["top"]  + data["height"]/2) / h
    data = data[["text","conf","x","y"]].reset_index(drop=True)
    # keep header-like tokens (short phrases). Still, we keep all; matching will pick correct ones.
    return data

def embed_texts(texts: List[str]) -> np.ndarray:
    if USE_ST:
        return ST_MODEL.encode(texts, normalize_embeddings=True)
    # fallback — fake embedding using itself; we will use fuzz later
    return np.array(texts, dtype=object)

def sem_similarity(a_vec, b_vec) -> float:
    if USE_ST:
        sim = float(np.dot(a_vec, b_vec))
        return max(0.0, min(1.0, sim))
    # fallback with RapidFuzz ratio (0..100)
    return fuzz.token_set_ratio(a_vec, b_vec) / 100.0

def pos_similarity(x1,y1,x2,y2) -> float:
    dist = math.sqrt((x1-x2)**2 + (y1-y2)**2)
    # max possible diag distance in normalized coords is sqrt(2)
    return max(0.0, 1.0 - dist / math.sqrt(2))

def decide_status(sim_sem, sim_pos, ocr_conf, en_found: bool) -> Tuple[int,str,str]:
    # Missing views:
    if not en_found:
        return 4, STATUS_DESCRIPTIONS[4], "Target label not found on English view"

    if np.isnan(ocr_conf) or ocr_conf < THRESH_OCR_CONF:
        return 0, STATUS_DESCRIPTIONS[0], "Low OCR confidence"

    if sim_sem >= THRESH_SEM_HIGH and sim_pos >= THRESH_SEM_MED:
        return 1, STATUS_DESCRIPTIONS[1], ""
    if sim_sem >= THRESH_SEM_MED:
        return 2, STATUS_DESCRIPTIONS[2], "Meaning differs / synonym or nuance"
    # Could be untranslated or very different
    return 3, STATUS_DESCRIPTIONS[3], "No good semantic match"

def best_match_for_label(gr_text, gr_xy, en_df, en_embeds, gr_embed):
    # Try all EN tokens, score by combined metric
    if en_df.empty:
        return None

    sem_sims = []
    for i, row in en_df.iterrows():
        if USE_ST:
            s = sem_similarity(gr_embed, en_embeds[i])
        else:
            s = sem_similarity(gr_text, row["text"])
        sem_sims.append(s)

    sem_sims = np.array(sem_sims)
    pos_sims = np.sqrt(
        (1 - ((en_df["x"].to_numpy() - gr_xy[0])**2 + (en_df["y"].to_numpy() - gr_xy[1])**2) / 2.0).clip(0,1)
    )  # quick smooth
    pos_sims = pos_sims.astype(float)

    final = WEIGHT_SEM*sem_sims + WEIGHT_POS*pos_sims
    j = int(final.argmax())
    return {
        "en_text": en_df.iloc[j]["text"],
        "en_conf": float(en_df.iloc[j]["conf"]),
        "en_x": float(en_df.iloc[j]["x"]),
        "en_y": float(en_df.iloc[j]["y"]),
        "sim_sem": float(sem_sims[j]),
        "sim_pos": float(pos_sims[j]),
        "final": float(final[j]),
    }

def build_report(trans_df: pd.DataFrame, gr_df: pd.DataFrame, en_df: pd.DataFrame) -> pd.DataFrame:
    # Pre-embed English tokens for speed
    en_embeds = embed_texts(en_df["text"].tolist())

    rows = []
    for _, tr in trans_df.iterrows():
        greek_term = tr["Greek"].strip()
        english_expected = tr["English"].strip()

        # Find Greek term location on Greek screenshot (best token)
        if gr_df.empty or greek_term == "":
            rows.append({
                "Greek": greek_term, "English(Expected)": english_expected,
                "English(Found)": "", "Semantic": 0.0, "Position": 0.0,
                "Greek_OCR_Conf": np.nan, "Greek_x": np.nan, "Greek_y": np.nan,
                "English_OCR_Conf": np.nan, "English_x": np.nan, "English_y": np.nan,
                "Status": 4, "Status Description": STATUS_DESCRIPTIONS[4],
                "Reason": "Greek label not detected on Greek screenshot"
            })
            continue

        # pick the Greek token with max similarity to target Greek text (in Greek)
        # This helps when OCR slightly breaks Greek words
        if USE_ST:
            gr_target_vec = embed_texts([greek_term])[0]
            gr_sims = []
            gr_embeds_list = embed_texts(gr_df["text"].tolist())
            for i in range(len(gr_df)):
                gr_sims.append(sem_similarity(gr_target_vec, gr_embeds_list[i]))
        else:
            gr_sims = [fuzz.token_set_ratio(greek_term, t)/100.0 for t in gr_df["text"]]

        gi = int(np.argmax(gr_sims)) if len(gr_sims) else None
        if gi is None or (len(gr_sims) and gr_sims[gi] < 0.60):
            # Greek token not confidently found on view
            rows.append({
                "Greek": greek_term, "English(Expected)": english_expected,
                "English(Found)": "", "Semantic": 0.0, "Position": 0.0,
                "Greek_OCR_Conf": np.nan, "Greek_x": np.nan, "Greek_y": np.nan,
                "English_OCR_Conf": np.nan, "English_x": np.nan, "English_y": np.nan,
                "Status": 4, "Status Description": STATUS_DESCRIPTIONS[4],
                "Reason": "Greek label not detected on Greek screenshot"
            })
            continue

        g_row = gr_df.iloc[gi]
        g_text, g_x, g_y, g_conf = g_row["text"], float(g_row["x"]), float(g_row["y"]), float(g_row["conf"])

        # Build Greek embedding from the detected OCR token (not from the dictionary string)
        gr_embed_for_match = embed_texts([g_text])[0]

        # Find best English match on English view
        match = best_match_for_label(g_text, (g_x, g_y), en_df, en_embeds, gr_embed_for_match)
        if match is None:
            status, desc, reason = decide_status(0.0, 0.0, g_conf, en_found=False)
            rows.append({
                "Greek": greek_term, "English(Expected)": english_expected,
                "English(Found)": "", "Semantic": 0.0, "Position": 0.0,
                "Greek_OCR_Conf": g_conf, "Greek_x": g_x, "Greek_y": g_y,
                "English_OCR_Conf": np.nan, "English_x": np.nan, "English_y": np.nan,
                "Status": status, "Status Description": desc, "Reason": reason
            })
            continue

        status, desc, reason = decide_status(match["sim_sem"], match["sim_pos"], g_conf, en_found=True)

        rows.append({
            "Greek": greek_term,
            "English(Expected)": english_expected,
            "English(Found)": match["en_text"],
            "Semantic": round(match["sim_sem"],3),
            "Position": round(match["sim_pos"],3),
            "Greek_OCR_Conf": round(g_conf,1),
            "Greek_x": round(g_x,3), "Greek_y": round(g_y,3),
            "English_OCR_Conf": round(match["en_conf"],1),
            "English_x": round(match["en_x"],3), "English_y": round(match["en_y"],3),
            "Status": status, "Status Description": desc, "Reason": reason
        })

    out = pd.DataFrame(rows)
    # Order: problems first
    cat = pd.CategoricalDtype([4,3,2,5,0,1], ordered=True)
    out["Status"] = out["Status"].astype(cat)
    out = out.sort_values(["Status","Greek"]).reset_index(drop=True)
    return out

# ---------- Run ----------
run = st.button("Run audit", type="primary", use_container_width=True)
if run:
    try:
        if not (xls_file and img_gr and img_en):
            st.error("Please upload: Excel + Greek screenshot + English screenshot.")
            st.stop()

        with st.spinner("Reading Excel…"):
            trans_df = load_excel_pairs(xls_file)

        with st.spinner("OCR: Greek view…"):
            gr_df = ocr_boxes(img_gr)
        with st.spinner("OCR: English view…"):
            en_df = ocr_boxes(img_en)

        if gr_df.empty:
            st.warning("No OCR tokens detected on the Greek screenshot.")
        if en_df.empty:
            st.warning("No OCR tokens detected on the English screenshot.")

        with st.spinner("Matching & scoring…"):
            report = build_report(trans_df, gr_df, en_df)

        st.success("Audit complete.")
        st.dataframe(report, use_container_width=True, height=420)

        # Download
        buf = io.BytesIO()
        report.to_excel(buf, index=False)
        buf.seek(0)
        st.download_button(
            "Download report (Excel)",
            data=buf,
            file_name="erp_translation_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Error: {e}")
