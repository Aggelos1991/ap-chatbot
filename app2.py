import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz

st.set_page_config(page_title="🤝 Vendor Reconciliation", layout="wide")
st.title("🧾 Universal Vendor Reconciliation App (AP Accurate)")

# ============================================================
# Column normalization (EN/ES/GR)
# ============================================================
def normalize_columns(df, source="ven"):
    colmap = {
        "vendor": ["supplier name", "vendor", "proveedor", "προμηθευτής"],
        "trn": ["tax id", "cif", "vat", "afm", "trn", "vat number", "tax number"],
        "invoice": ["invoice number", "alt document", "invoice", "παραστατικό", "factura"],
        "description": ["description", "descripción", "περιγραφή"],
        "debit": ["debit", "debe", "χρέωση"],
        "credit": ["credit", "haber", "πίστωση"],
        "amount": ["amount", "importe", "valor"],
        "balance": ["balance", "saldo", "υπόλοιπο"],
        "date": ["date", "fecha", "ημερομηνία"]
    }

    rename_map = {}
    for key, variants in colmap.items():
        for col in df.columns:
            col_low = str(col).strip().lower()
            if any(v in col_low for v in variants):
                rename_map[col] = f"{key}_{source}"
                break

    df = df.rename(columns=rename_map)
    return df


# ============================================================
# Matching Logic (Invoices + CN only, accurate AP treatment)
# ============================================================
def match_invoices(erp_df, ven_df):
    matched, erp_missing = [], []
    ven_missing = ven_df.copy().reset_index(drop=True)

    # Filter ERP to same vendor TRN
    vendor_trn = str(ven_df["trn_ven"].iloc[0]) if "trn_ven" in ven_df.columns else None
    if vendor_trn:
        erp_df = erp_df[erp_df["trn_erp"] == vendor_trn]

    for _, erp_row in erp_df.iterrows():
        erp_trn = str(erp_row.get("trn_erp", "")).strip()
        erp_invoice = str(erp_row.get("invoice_erp", "")).strip()

        # ✅ Use Amount or Credit for invoice value, ignore Balance
        erp_amount = float(
            erp_row.get("amount_erp", erp_row.get("credit_erp", 0)) or 0
        )

        ven_subset = ven_missing[ven_missing["trn_ven"] == erp_trn]
        found = False

        for _, ven_row in ven_subset.iterrows():
            ven_invoice = str(ven_row.get("invoice_ven", "")).strip()
            desc = str(ven_row.get("description_ven", "")).lower()

            debit_val = float(ven_row.get("debit_ven", 0) or 0)
            credit_val = float(ven_row.get("credit_ven", 0) or 0)

            # ✅ Vendor side: invoice = Debit; CN = Credit if "abono"
            if "abono" in desc or "credit" in desc:
                ven_amount = credit_val
            elif any(w in desc for w in ["pago", "transferencia", "payment", "πληρωμή"]):
                continue  # 🚫 Skip payments
            else:
                ven_amount = debit_val

            # --- Flexible invoice matching
            if (
                erp_invoice[-4:] in ven_invoice
                or ven_invoice[-4:] in erp_invoice
                or fuzz.ratio(erp_invoice, ven_invoice) > 75
            ):
                diff = round(erp_amount - ven_amount, 2)
                status = "Match" if diff == 0 else "Difference"

                matched.append({
                    "Vendor/Supplier": erp_row.get("vendor_erp", ""),
                    "TRN/AFM": erp_trn,
                    "ERP Invoice": erp_invoice,
                    "Vendor Invoice": ven_invoice,
                    "ERP Amount": erp_amount,
                    "Vendor Amount": ven_amount,
                    "Difference": diff,
                    "Status": status,
                    "Description": desc
                })

                if ven_row.name in ven_missing.index:
                    ven_missing = ven_missing.drop(index=ven_row.name)
                found = True
                break

        if not found:
            erp_missing.append(erp_row)

    return pd.DataFrame(matched), pd.DataFrame(erp_missing), ven_missing.reset_index(drop=True)


# ============================================================
# Streamlit Interface
# ============================================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = pd.read_excel(uploaded_erp)
    ven_df = pd.read_excel(uploaded_vendor)

    erp_df = normalize_columns(erp_df, source="erp")
    ven_df = normalize_columns(ven_df, source="ven")

    with st.spinner("🔍 Reconciling invoices and credit notes..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    st.success("✅ Reconciliation complete!")

    st.subheader("📊 Matched / Differences")
    st.dataframe(matched)

    st.subheader("❌ In ERP but Missing in Vendor")
    st.dataframe(erp_missing)

    st.subheader("❌ In Vendor but Missing in ERP")
    st.dataframe(ven_missing)

    st.download_button(
        "⬇️ Download Matched Results",
        matched.to_csv(index=False).encode("utf-8"),
        "matched_results.csv",
        "text/csv"
    )

else:
    st.info("Please upload both ERP Export and Vendor Statement files to begin.")
