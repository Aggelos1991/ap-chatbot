import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="DataFalcon — FINAL VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL VERSION (TABLE EXTRACTION)")

uploaded = st.file_uploader("Upload Ledger PDF", type=["pdf"])

def clean_amount(v):
    if not v or v.strip() == "":
        return 0.0
    v = v.replace(".", "").replace(",", ".")
    try:
        return float(v)
    except:
        return 0.0

if uploaded:
    rows = []

    with pdfplumber.open(uploaded) as pdf:
        for page in pdf.pages:

            # TABLE EXTRACTION — THIS IS THE REAL FIX
            table = page.extract_table({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 5,
                "intersection_x_tolerance": 5
            })

            if not table:
                continue

            # FIRST ROW = HEADERS → SKIP
            for row in table[1:]:

                if all(v is None or v.strip() == "" for v in row):
                    continue

                date = row[0]
                asiento = row[1]
                documento = row[2]
                libro = row[3]
                descripcion = row[4]
                referencia = row[5]
                fvalor = row[6]
                debe = clean_amount(row[7])
                haber = clean_amount(row[8])
                saldo = clean_amount(row[9])

                rows.append({
                    "Date": date,
                    "Asiento": asiento,
                    "Documento": documento,
                    "Libro": libro,
                    "Descripcion": descripcion,
                    "Referencia": referencia,
                    "F_Valor": fvalor,
                    "Debe": debe,
                    "Haber": haber,
                    "Saldo": saldo
                })

    df = pd.DataFrame(rows)

    st.success(f"EXTRACTED {len(df)} REAL LEDGER ROWS ✔")
    st.dataframe(df, use_container_width=True)

    # DOWNLOAD
    buf = BytesIO()
    df.to_excel(buf, index=False)
    st.download_button(
        "Download Excel",
        buf.getvalue(),
        "DataFalcon_FINAL.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
