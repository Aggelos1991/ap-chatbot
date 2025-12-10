import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="DataFalcon — FINAL SAFE VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL SAFE VERSION (NO CRASH)")

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

            table = page.extract_table({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 5,
                "intersection_x_tolerance": 5
            })

            if not table:
                continue

            # Skip header row
            for row in table[1:]:
                # FIX: pad row so row[0..9] always exist
                if row is None:
                    continue

                # Make row always length 10
                while len(row) < 10:
                    row.append("")

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

                # Skip completely empty rows
                if not date and not descripcion and debe == 0 and haber == 0:
                    continue

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

    st.success(f"EXTRACTED {len(df)} LEDGER ROWS ✔ (SAFE MODE)")
    st.dataframe(df, use_container_width=True)

    # Export
    buf = BytesIO()
    df.to_excel(buf, index=False)

    st.download_button(
        "Download Excel",
        buf.getvalue(),
        "DataFalcon_FINAL.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
