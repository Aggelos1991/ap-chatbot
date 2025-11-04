# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FINAL • Payments Fix • Tier de-dup)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows, get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher

# ==================== PAGE CONFIG & CSS ======================
st.set_page_config(page_title="ReconRaptor — Vendor Reconciliation", layout="wide")
st.markdown("""
<style>
.big-title {
    font-size: 3rem !important;
    font-weight: 700;
    text-align: center;
    background: linear-gradient(90deg, #1E88E5, #42A5F5);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 1rem;
}
.section-title {
    font-size: 1.8rem !important;
    font-weight: 600;
    color: #1565C0;
    border-bottom: 2px solid #42A5F5;
    padding-bottom: 0.5rem;
    margin-top: 2rem;
}
.metric-container {
    padding: 1.2rem !important;
    border-radius: 12px !important;
    margin-bottom: 1rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
}
.perfect-match   {background:#2E7D32;color:#fff;font-weight:bold;}
.difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
.tier2-match     {background:#26A69A;color:#fff;font-weight:bold;}
.tier3-match     {background:#7E57C2;color:#fff;font-weight:bold;}
.missing-erp     {background:#C62828;color:#fff;font-weight:bold;}
.missing-vendor  {background:#AD1457;color:#fff;font-weight:bold;}
.payment-match   {background:#004D40;color:#fff;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a,b): return SequenceMatcher(None,str(a),str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip()=="":
        return 0.0
    s=re.sub(r"[^\d,.\-]","",str(v).strip())
    if s.count(",")==1 and s.count(".")==1:
        if s.find(",")>s.find("."): s=s.replace(".","").replace(",",".")
        else: s=s.replace(",", "")
    elif s.count(",")==1: s=s.replace(",",".")
    elif s.count(".")>1: s=s.replace(".","",s.count(".")-1)
    try: return float(s)
    except: return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip()=="": return ""
    s=str(v).strip().replace(".","/").replace("-","/").replace(",","/")
    for fmt in ["%d/%m/%Y","%m/%d/%Y","%Y/%m/%d","%d/%m/%y","%Y.%m.%d"]:
        try:
            d=pd.to_datetime(s,format=fmt,errors="coerce")
            if not pd.isna(d): return d.strftime("%Y-%m-%d")
        except: continue
    d=pd.to_datetime(s,errors="coerce",dayfirst=True)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s=str(v).lower().strip()
    s=re.sub(r"[^a-z0-9]","",s)
    s=re.sub(r"^0+","",s)
    return s or "0"

def normalize_columns(df,tag):
    mapping={
        "invoice":["invoice","factura","fact","num","document","ref"],
        "credit":["credit","haber","credito","abono"],
        "debit":["debit","debe","cargo","importe","amount","valor","total"],
        "reason":["reason","motivo","concepto","descripcion","detalle"],
        "date":["date","fecha","data"]
    }
    rename_map={}
    cols_lower={c: str(c).lower() for c in df.columns}
    for key,aliases in mapping.items():
        for col,low in cols_lower.items():
            if any(a in low for a in aliases): rename_map[col]=f"{key}_{tag}"
    out=df.rename(columns=rename_map)
    for c in ["debit","credit"]:
        col=f"{c}_{tag}"
        if col not in out.columns: out[col]=0.0
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"]=out[f"date_{tag}"].apply(normalize_date)
    return out

def style(df,css): return df.style.apply(lambda _: [css]*len(_),axis=1)

# ==================== MATCHING CORE (unchanged logic) ==========================
def match_invoices(erp_df,ven_df):
    def doc_type(row,tag):
        txt=(str(row.get(f"reason_{tag}",""))+" "+str(row.get(f"invoice_{tag}",""))).lower()
        debit=normalize_number(row.get(f"debit_{tag}",0))
        credit=normalize_number(row.get(f"credit_{tag}",0))
        pay_kw=["πληρωμ","payment","remittance","transferencia","pago","cobro"]
        if any(k in txt for k in pay_kw): return "IGNORE"
        if any(k in txt for k in ["credit","abono","cn","nota"]): return "CN"
        if debit!=0 or credit!=0: return "INV"
        return "UNKNOWN"

    for df,tag in [(erp_df,"erp"),(ven_df,"ven")]:
        df["__type"]=df.apply(lambda r: doc_type(r,tag),axis=1)
        df["__amt"]=df.apply(lambda r: abs(normalize_number(r.get(f"debit_{tag}",0))-normalize_number(r.get(f"credit_{tag}",0))),axis=1)

    erp_use=erp_df[erp_df["__type"]!="IGNORE"].copy()
    ven_use=ven_df[ven_df["__type"]!="IGNORE"].copy()

    matched=[]
    used=set()
    for ei,e in erp_use.iterrows():
        for vi,v in ven_use.iterrows():
            if vi in used: continue
            if str(e.get("invoice_erp","")).strip()==str(v.get("invoice_ven","")).strip():
                diff=abs(e["__amt"]-v["__amt"])
                matched.append({
                    "ERP Invoice":e["invoice_erp"],"Vendor Invoice":v["invoice_ven"],
                    "ERP Amount":e["__amt"],"Vendor Amount":v["__amt"],
                    "Difference":round(diff,2),
                    "Status":"Perfect Match" if diff<=0.01 else "Difference Match"
                })
                used.add(vi)
                break
    m=pd.DataFrame(matched)
    miss_erp=erp_use[~erp_use["invoice_erp"].isin(m["ERP Invoice"] if not m.empty else [])]
    miss_ven=ven_use[~ven_use["invoice_ven"].isin(m["Vendor Invoice"] if not m.empty else [])]
    miss_erp=miss_erp.rename(columns={"invoice_erp":"Invoice","__amt":"Amount","date_erp":"Date"})
    miss_ven=miss_ven.rename(columns={"invoice_ven":"Invoice","__amt":"Amount","date_ven":"Date"})
    return m,miss_erp,miss_ven

# ==================== PAYMENT DETECTION FIX ==========================
def extract_payments(erp_df,ven_df):
    pay_kw=["πληρωμή","payment","remittance","transferencia","trf","remesa","pago","deposit","έμβασμα","paid","cobro"]
    excl_kw=["invoice of expenses","reclass","adjustment","διόρθωση"]

    def is_payment(row,tag):
        txt=(str(row.get(f"reason_{tag}",""))+" "+str(row.get(f"invoice_{tag}",""))).lower()
        return any(k in txt for k in pay_kw) and not any(b in txt for b in excl_kw)

    erp_pay=erp_df[erp_df.apply(lambda r:is_payment(r,"erp"),axis=1)].copy()
    ven_pay=ven_df[ven_df.apply(lambda r:is_payment(r,"ven"),axis=1)].copy()

    for df,tag in [(erp_pay,"erp"),(ven_pay,"ven")]:
        if not df.empty:
            df["Debit"]=df[f"debit_{tag}"].apply(normalize_number)
            df["Credit"]=df[f"credit_{tag}"].apply(normalize_number)
            # FIX: choose whichever side is nonzero if both exist
            df["Amount"]=df.apply(lambda r: r["Debit"] if abs(r["Debit"])>0 else abs(r["Credit"]),axis=1)
            df["Amount"]=df["Amount"].round(2)

    matched=[]
    used=set()
    for ei,e in erp_pay.iterrows():
        for vi,v in ven_pay.iterrows():
            if vi in used: continue
            if abs(e["Amount"]-v["Amount"])<=0.05:
                matched.append({
                    "ERP Reason":e.get("reason_erp",""),
                    "Vendor Reason":v.get("reason_ven",""),
                    "ERP Amount":round(e["Amount"],2),
                    "Vendor Amount":round(v["Amount"],2),
                    "Difference":round(abs(e["Amount"]-v["Amount"]),2)
                })
                used.add(vi)
                break
    return erp_pay,ven_pay,pd.DataFrame(matched)

# ==================== EXPORT + UI ==========================
def export_excel(m1,m2):
    wb=Workbook();ws=wb.active;ws.title="Missing"
    ws.append(["Missing in ERP"]);[ws.append(r) for r in dataframe_to_rows(m2,index=False,header=True)]
    ws.append([]);ws.append(["Missing in Vendor"]);[ws.append(r) for r in dataframe_to_rows(m1,index=False,header=True)]
    for c in ws.columns:
        ws.column_dimensions[get_column_letter(c[0].column)].width=max(len(str(x.value)) for x in c)+2
    buf=BytesIO();wb.save(buf);buf.seek(0);return buf

st.markdown("### Upload Your Files")
u1=st.file_uploader("ERP Export",type=["xlsx"],key="erp")
u2=st.file_uploader("Vendor Statement",type=["xlsx"],key="vendor")

if u1 and u2:
    erp_raw=pd.read_excel(u1,dtype=str)
    ven_raw=pd.read_excel(u2,dtype=str)
    erp_df=normalize_columns(erp_raw,"erp")
    ven_df=normalize_columns(ven_raw,"ven")

    st.write("ERP cols:",list(erp_df.columns))
    st.write("Vendor cols:",list(ven_df.columns))

    with st.spinner("Reconciling..."):
        tier1,miss_erp,miss_ven=match_invoices(erp_df,ven_df)
        erp_pay,ven_pay,pay_match=extract_payments(erp_df,ven_df)

    st.success("✅ Reconciliation Complete")
    st.subheader("Tier-1 Matches")
    st.dataframe(tier1,use_container_width=True)

    st.subheader("ERP Payments (fixed)")
    if not erp_pay.empty:
        st.dataframe(erp_pay[["reason_erp","Debit","Credit","Amount"]],use_container_width=True)
        st.write("**Total ERP Payments:**",erp_pay["Amount"].sum())
    else:
        st.info("No ERP payments found.")

    st.subheader("Vendor Payments")
    if not ven_pay.empty:
        st.dataframe(ven_pay[["reason_ven","Debit","Credit","Amount"]],use_container_width=True)
        st.write("**Total Vendor Payments:**",ven_pay["Amount"].sum())
    else:
        st.info("No Vendor payments found.")

    if not pay_match.empty:
        st.subheader("Matched Payments")
        st.dataframe(pay_match,use_container_width=True)

    buf=export_excel(miss_erp,miss_ven)
    st.download_button("Download Excel Report",data=buf,file_name="ReconRaptor_Report.xlsx")
