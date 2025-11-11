# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FINAL FULL STABLE VERSION)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows, get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher
import numpy as np

# ==================== PAGE CONFIG & CSS ======================
st.set_page_config(page_title="ReconRaptor — Vendor Reconciliation", layout="wide")
st.markdown("""
<style>
.big-title {font-size:3rem!important;font-weight:700;text-align:center;
background:linear-gradient(90deg,#1E88E5,#42A5F5);
-webkit-background-clip:text;-webkit-text-fill-color:transparent;
margin-bottom:1rem;}
.section-title {font-size:1.8rem!important;font-weight:600;color:#1565C0;
border-bottom:2px solid #42A5F5;padding-bottom:0.5rem;margin-top:2rem;}
.metric-container {padding:1.2rem;border-radius:12px;margin-bottom:1rem;
box-shadow:0 4px 6px rgba(0,0,0,0.1);}
.perfect-match{background:#2E7D32;color:#fff;font-weight:bold;}
.difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
.tier2-match{background:#26A69A;color:#fff;font-weight:bold;}
.tier3-match{background:#7E57C2;color:#fff;font-weight:bold;}
.missing-erp{background:#C62828;color:#fff;font-weight:bold;}
.missing-vendor{background:#AD1457;color:#fff;font-weight:bold;}
.payment-match{background:#004D40;color:#fff;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a,b): return SequenceMatcher(None,str(a),str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip()=="":
        return 0.0
    s = re.sub(r"[^\d,.\-]","",str(v).strip())
    if s.count(",")==1 and s.count(".")==1:
        if s.find(",")>s.find("."):
            s=s.replace(".","").replace(",",".")
        else:
            s=s.replace(",","")
    elif s.count(",")==1:
        s=s.replace(",",".")
    elif s.count(".")>1:
        s=s.replace(".","",s.count(".")-1)
    try:
        return float(s)
    except:
        return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip()=="":
        return ""
    s=str(v).strip().replace(".","/").replace("-","/").replace(",","/")
    d=pd.to_datetime(s,errors="coerce",dayfirst=True)
    if pd.isna(d):
        d=pd.to_datetime(s,errors="coerce",dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s=str(v).strip().lower()
    parts=re.split(r"[-_.\s]",s)
    for p in reversed(parts):
        if re.fullmatch(r"\d{1,}",p) and not re.fullmatch(r"20[0-3]\d",p):
            s=p.lstrip("0"); break
    s=re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|apd|vs)\W*","",s)
    s=re.sub(r"20\d{2}","",s)
    s=re.sub(r"[^a-z0-9]","",s)
    s=re.sub(r"^0+","",s)
    s=re.sub(r"[^\d]","",s)
    return s or "0"

def normalize_columns(df,tag):
    mapping={
        "invoice":["invoice","invoice number","inv no","factura","fact","numero","document","ref","alternative document","alt document","alt. document"],
        "credit":["credit","haber","credito","abono"],
        "debit":["debit","debe","cargo","importe","amount","valor","total","charge"],
        "reason":["reason","motivo","concepto","descripcion","detalle"],
        "date":["date","fecha","data","issue date","posting date"]
    }
    rename_map={}
    cols_lower={c:str(c).strip().lower() for c in df.columns}
    for key,aliases in mapping.items():
        for col,low in cols_lower.items():
            if any(a in low for a in aliases): rename_map[col]=f"{key}_{tag}"
    out=df.rename(columns=rename_map)
    for req in ["debit","credit"]:
        c=f"{req}_{tag}"
        if c not in out.columns: out[c]=0.0
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"]=out[f"date_{tag}"].apply(normalize_date)
    return out

def style(df,css): return df.style.apply(lambda _:[css]*len(_),axis=1)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df,ven_df):
    def doc_type(row,tag):
        txt=(str(row.get(f"reason_{tag}",""))+" "+str(row.get(f"invoice_{tag}",""))).lower()
        debit,credit=normalize_number(row.get(f"debit_{tag}",0)),normalize_number(row.get(f"credit_{tag}",0))
        if any(k in txt for k in ["payment","remittance","transferencia","pago","paid"]): return "IGNORE"
        if any(k in txt for k in ["credit","nota","abono","cn"]): return "CN"
        if any(k in txt for k in ["factura","invoice","inv"]) or debit>0 or credit>0: return "INV"
        return "UNKNOWN"
    erp_df["__type"]=erp_df.apply(lambda r:doc_type(r,"erp"),axis=1)
    ven_df["__type"]=ven_df.apply(lambda r:doc_type(r,"ven"),axis=1)

    def compute_amt(row,tag):
        debit,credit=normalize_number(row.get(f"debit_{tag}",0)),normalize_number(row.get(f"credit_{tag}",0))
        if debit==0 and credit==0:
            for col in row.index:
                if isinstance(col,str) and "charge" in col.lower():
                    v=normalize_number(row.get(col,0))
                    if v!=0: return round(abs(v),2)
        return round(abs(debit-credit),2)

    erp_df["__amt"]=erp_df.apply(lambda r:compute_amt(r,"erp"),axis=1)
    ven_df["__amt"]=ven_df.apply(lambda r:compute_amt(r,"ven"),axis=1)

    matched,used_v=[],set()
    for ei,e in erp_df[erp_df["__type"]!="IGNORE"].iterrows():
        e_inv=str(e.get("invoice_erp","")).strip()
        e_amt=round(float(e.get("__amt",0.0)),2)
        for vi,v in ven_df[ven_df["__type"]!="IGNORE"].iterrows():
            if vi in used_v: continue
            v_inv=str(v.get("invoice_ven","")).strip()
            v_amt=round(float(v.get("__amt",0.0)),2)
            if e_inv==v_inv:
                diff=abs(e_amt-v_amt)
                matched.append({
                    "ERP Invoice":e_inv,"Vendor Invoice":v_inv,
                    "ERP Amount":e_amt,"Vendor Amount":v_amt,
                    "Difference":round(diff,2),
                    "Status":"Perfect Match" if diff<=0.01 else "Difference Match"
                })
                used_v.add(vi); break
    matched_df=pd.DataFrame(matched)
    miss_erp=erp_df[~erp_df["invoice_erp"].isin(matched_df["ERP Invoice"] if not matched_df.empty else [])]
    miss_ven=ven_df[~ven_df["invoice_ven"].isin(matched_df["Vendor Invoice"] if not matched_df.empty else [])]
    miss_erp=miss_erp.rename(columns={"invoice_erp":"Invoice","__amt":"Amount","date_erp":"Date"})
    miss_ven=miss_ven.rename(columns={"invoice_ven":"Invoice","__amt":"Amount","date_ven":"Date"})
    keep=["Invoice","Amount","Date"]
    return matched_df,miss_erp[[c for c in keep if c in miss_erp.columns]],miss_ven[[c for c in keep if c in miss_ven.columns]]

# ==================== EXCEL EXPORT ==========================
def export_excel(miss_erp,miss_ven):
    wb=Workbook(); wb.remove(wb.active)
    ws=wb.create_sheet("Missing")
    cur=1
    def hdr(ws,row,color):
        for c in ws[row]:
            c.fill=PatternFill(start_color=color,end_color=color,fill_type="solid")
            c.font=Font(color="FFFFFF",bold=True)
            c.alignment=Alignment(horizontal="center",vertical="center")
    if not miss_ven.empty:
        ws.merge_cells(start_row=cur,start_column=1,end_row=cur,end_column=max(3,miss_ven.shape[1]))
        ws.cell(cur,1,"Missing in ERP").font=Font(bold=True,size=14); cur+=2
        for r in dataframe_to_rows(miss_ven,index=False,header=True): ws.append(r)
        hdr(ws,cur,"C62828"); cur=ws.max_row+3
    if not miss_erp.empty:
        ws.merge_cells(start_row=cur,start_column=1,end_row=cur,end_column=max(3,miss_erp.shape[1]))
        ws.cell(cur,1,"Missing in Vendor").font=Font(bold=True,size=14); cur+=2
        for r in dataframe_to_rows(miss_erp,index=False,header=True): ws.append(r)
        hdr(ws,cur,"AD1457")
    for col in ws.columns:
        max_len=max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width=max_len+3
    buf=BytesIO(); wb.save(buf); buf.seek(0); return buf

# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp=st.file_uploader("ERP Export (Excel)",type=["xlsx"],key="erp")
uploaded_vendor=st.file_uploader("Vendor Statement (Excel)",type=["xlsx"],key="vendor")

if uploaded_erp and uploaded_vendor:
    try:
        # --- universal cleaner ---
        def clean_excel(df):
            df.columns=[str(c).strip().lower() for c in df.columns]
            df=df.applymap(lambda x:str(x).strip() if pd.notna(x) else x)
            for c in df.columns:
                if any(k in c for k in ["debit","credit","importe","valor","amount","total","charge"]):
                    df[c]=(df[c].astype(str)
                           .str.replace(".","",regex=False)
                           .str.replace(",",".",regex=False))
            return df

        erp_raw=clean_excel(pd.read_excel(uploaded_erp,dtype=str))
        ven_raw=clean_excel(pd.read_excel(uploaded_vendor,dtype=str))
        erp_df=normalize_columns(erp_raw,"erp")
        ven_df=normalize_columns(ven_raw,"ven")

        with st.spinner("Analyzing invoices..."):
            tier1,miss_erp,miss_ven=match_invoices(erp_df,ven_df)

            def safe_series_to_str(obj):
                if isinstance(obj,pd.Series): return obj.astype(str)
                elif isinstance(obj,list): return pd.Series(obj,dtype=str)
                elif isinstance(obj,str): return pd.Series([obj],dtype=str)
                else: return pd.Series([],dtype=str)

            used_erp_inv=set(); used_ven_inv=set()
            if isinstance(tier1,pd.DataFrame) and not tier1.empty:
                if "ERP Invoice" in tier1.columns:
                    used_erp_inv=set(safe_series_to_str(tier1["ERP Invoice"]))
                if "Vendor Invoice" in tier1.columns:
                    used_ven_inv=set(safe_series_to_str(tier1["Vendor Invoice"]))

        st.success("Reconciliation Complete!")
        st.markdown('<h2 class="section-title">Results</h2>', unsafe_allow_html=True)
        st.dataframe(tier1,use_container_width=True)
        st.markdown('<h2 class="section-title">Missing</h2>', unsafe_allow_html=True)
        st.write("**Missing in ERP**",miss_ven)
        st.write("**Missing in Vendor**",miss_erp)

        excel_buf=export_excel(miss_erp,miss_ven)
        st.download_button("Download Excel Report",excel_buf,"ReconRaptor_Report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that your files contain columns like: invoice, debit/credit, date, reason.")
