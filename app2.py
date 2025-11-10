# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FINAL • All Tiers Clean • FIXED)
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
st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a,b): return SequenceMatcher(None,str(a),str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip()=="": return 0.0
    s=re.sub(r"[^\d,.\-]","",str(v).strip())
    if s.count(",")==1 and s.count(".")==1:
        if s.find(",")>s.find("."): s=s.replace(".","").replace(",",".")
        else: s=s.replace(",", "")
    elif s.count(",")==1: s=s.replace(",",".")
    elif s.count(".")>1: s=s.replace(".", "", s.count(".")-1)
    try: return float(s)
    except: return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip()=="": return ""
    s=str(v).strip().replace(".","/").replace("-","/").replace(",", "/")
    for fmt in ["%d/%m/%Y","%d-%m-%Y","%m/%d/%Y","%Y/%m/%d","%d/%m/%y","%m/%d/%y","%Y-%m-%d"]:
        try:
            d=pd.to_datetime(s,format=fmt,errors="coerce")
            if not pd.isna(d): return d.strftime("%Y-%m-%d")
        except: continue
    d=pd.to_datetime(s,errors="coerce",dayfirst=True)
    if pd.isna(d): d=pd.to_datetime(s,errors="coerce",dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s=str(v).strip().lower()
    s=re.sub(r"[^a-z0-9]","",s)
    s=re.sub(r"^0+","",s)
    return s or "0"

def normalize_columns(df,tag):
    mapping={"invoice":["invoice","factura","document","ref"],
             "credit":["credit","haber","abono"],
             "debit":["debit","debe","importe","amount","valor","total"],
             "reason":["reason","motivo","descripcion","detalle"],
             "date":["date","fecha","data","issue","posting"]}
    rename={}
    low={c:str(c).lower() for c in df.columns}
    for key,alias in mapping.items():
        for c,l in low.items():
            if any(a in l for a in alias): rename[c]=f"{key}_{tag}"
    out=df.rename(columns=rename)
    for k in ["debit","credit"]:
        if f"{k}_{tag}" not in out.columns: out[f"{k}_{tag}"]=0.0
    if f"date_{tag}" in out.columns: out[f"date_{tag}"]=out[f"date_{tag}"].apply(normalize_date)
    return out

def style(df,css): return df.style.apply(lambda _: [css]*len(_),axis=1)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df,ven_df):
    erp_df["__amt"]=erp_df.apply(lambda r:abs(normalize_number(r.get("debit_erp",0))-normalize_number(r.get("credit_erp",0))),axis=1)
    ven_df["__amt"]=ven_df.apply(lambda r:abs(normalize_number(r.get("debit_ven",0))-normalize_number(r.get("credit_ven",0))),axis=1)
    matched=[];used=set()
    for i,e in erp_df.iterrows():
        e_inv=str(e.get("invoice_erp","")).strip(); e_amt=round(e["__amt"],2)
        for j,v in ven_df.iterrows():
            if j in used: continue
            v_inv=str(v.get("invoice_ven","")).strip(); v_amt=round(v["__amt"],2)
            if e_inv==v_inv:
                diff=abs(e_amt-v_amt)
                matched.append({"ERP Invoice":e_inv,"Vendor Invoice":v_inv,
                                "ERP Amount":e_amt,"Vendor Amount":v_amt,
                                "Difference":round(diff,2),
                                "Status":"Perfect Match" if diff<=0.01 else "Difference Match"})
                used.add(j);break
    matched_df=pd.DataFrame(matched)
    erp_df["__inv_norm"]=erp_df["invoice_erp"].apply(clean_invoice_code)
    ven_df["__inv_norm"]=ven_df["invoice_ven"].apply(clean_invoice_code)
    miss_erp=erp_df[~erp_df["__inv_norm"].isin(matched_df["ERP Invoice"].apply(clean_invoice_code))]
    miss_ven=ven_df[~ven_df["__inv_norm"].isin(matched_df["Vendor Invoice"].apply(clean_invoice_code))]
    miss_erp=miss_erp.rename(columns={"invoice_erp":"Invoice","__amt":"Amount","date_erp":"Date"})
    miss_ven=miss_ven.rename(columns={"invoice_ven":"Invoice","__amt":"Amount","date_ven":"Date"})
    keep=["Invoice","Amount","Date"]
    miss_erp=miss_erp[[c for c in keep if c in miss_erp.columns]].reset_index(drop=True)
    miss_ven=miss_ven[[c for c in keep if c in miss_ven.columns]].reset_index(drop=True)
    return matched_df,miss_erp,miss_ven

def tier2_match(erp_miss,ven_miss):
    if erp_miss.empty or ven_miss.empty: return pd.DataFrame(),set(),set(),erp_miss,ven_miss
    matches=[];used_e=set();used_v=set()
    for ei,er in erp_miss.iterrows():
        e_inv=str(er["Invoice"]); e_amt=round(float(er["Amount"]),2); e_code=clean_invoice_code(e_inv)
        for vi,vr in ven_miss.iterrows():
            if vi in used_v: continue
            v_inv=str(vr["Invoice"]); v_amt=round(float(vr["Amount"]),2); v_code=clean_invoice_code(v_inv)
            diff=abs(e_amt-v_amt); sim=fuzzy_ratio(e_code,v_code)
            if diff<=1.00 and sim>=0.70:
                matches.append({"ERP Invoice":e_inv,"Vendor Invoice":v_inv,
                                "ERP Amount":e_amt,"Vendor Amount":v_amt,
                                "Difference":round(diff,2),
                                "Fuzzy Score":round(sim,2),"Match Type":"Tier-2"})
                used_e.add(ei);used_v.add(vi);break
    mdf=pd.DataFrame(matches)
    rem_e=erp_miss[~erp_miss.index.isin(used_e)].copy()
    rem_v=ven_miss[~ven_miss.index.isin(used_v)].copy()
    return mdf,used_e,used_v,rem_e,rem_v

def tier3_match(erp_miss,ven_miss):
    if erp_miss.empty or ven_miss.empty: return pd.DataFrame(),set(),set(),erp_miss,ven_miss
    matches=[];used_e=set();used_v=set()
    for ei,er in erp_miss.iterrows():
        e_inv=str(er["Invoice"]); e_amt=round(float(er["Amount"]),2)
        e_date=normalize_date(er.get("Date","")); e_code=clean_invoice_code(e_inv)
        for vi,vr in ven_miss.iterrows():
            v_inv=str(vr["Invoice"]); v_amt=round(float(vr["Amount"]),2)
            v_date=normalize_date(vr.get("Date","")); v_code=clean_invoice_code(v_inv)
            sim=fuzzy_ratio(e_code,v_code)
            if e_date and e_date==v_date and sim>=0.75:
                diff=abs(e_amt-v_amt)
                matches.append({"ERP Invoice":e_inv,"Vendor Invoice":v_inv,
                                "ERP Amount":e_amt,"Vendor Amount":v_amt,
                                "Difference":round(diff,2),
                                "Fuzzy Score":round(sim,2),"Date":e_date,"Match Type":"Tier-3"})
                used_e.add(ei);used_v.add(vi);break
    mdf=pd.DataFrame(matches)
    rem_e=erp_miss[~erp_miss.index.isin(used_e)].copy()
    rem_v=ven_miss[~ven_miss.index.isin(used_v)].copy()
    return mdf,used_e,used_v,rem_e,rem_v

# ==================== PAYMENTS ==========================
def extract_payments(erp_df,ven_df):
    pay_kw=["πληρωμή","payment","remittance","bank transfer","transferencia","trf","remesa","pago","deposit","μεταφορά","έμβασμα","εξόφληση","pagado","paid","cobro"]
    def is_pay(row,tag):
        txt=(str(row.get(f"reason_{tag}",""))+str(row.get(f"invoice_{tag}",""))).lower()
        return any(k in txt for k in pay_kw)
    erp_pay=erp_df[erp_df.apply(lambda r:is_pay(r,"erp"),axis=1)].copy()
    ven_pay=ven_df[ven_df.apply(lambda r:is_pay(r,"ven"),axis=1)].copy()
    def compute_amount(df,tag):
        if df.empty: return df
        df["Debit"]=df[f"debit_{tag}"].apply(normalize_number)
        df["Credit"]=df[f"credit_{tag}"].apply(normalize_number)
        df["Amount"]=(df["Debit"]-df["Credit"]).abs().round(2)
        return df
    erp_pay=compute_amount(erp_pay,"erp"); ven_pay=compute_amount(ven_pay,"ven")
    matched=[];used=set()
    for _,e in erp_pay.iterrows():
        for vi,v in ven_pay.iterrows():
            if vi in used: continue
            if abs(e["Amount"]-v["Amount"])<=0.05:
                matched.append({"ERP Reason":e.get("reason_erp",""),"Vendor Reason":v.get("reason_ven",""),
                                "ERP Amount":e["Amount"],"Vendor Amount":v["Amount"],
                                "Difference":round(abs(e["Amount"]-v["Amount"]),2)})
                used.add(vi);break
    return erp_pay,ven_pay,pd.DataFrame(matched)

# ==================== EXPORT ==========================
def export_excel(miss_erp,miss_ven):
    wb=Workbook();wb.remove(wb.active);ws=wb.create_sheet("Missing")
    def hdr(ws,row,color):
        for c in ws[row]:
            c.fill=PatternFill(start_color=color,end_color=color,fill_type="solid")
            c.font=Font(color="FFFFFF",bold=True)
            c.alignment=Alignment(horizontal="center",vertical="center")
    cur=1
    if not miss_ven.empty:
        ws.merge_cells(start_row=cur,start_column=1,end_row=cur,end_column=max(3,miss_ven.shape[1]))
        ws.cell(cur,1,"Missing in ERP").font=Font(bold=True,size=14);cur+=2
        for r in dataframe_to_rows(miss_ven,index=False,header=True): ws.append(r)
        hdr(ws,cur,"C62828");cur=ws.max_row+3
    if not miss_erp.empty:
        ws.merge_cells(start_row=cur,start_column=1,end_row=cur,end_column=max(3,miss_erp.shape[1]))
        ws.cell(cur,1,"Missing in Vendor").font=Font(bold=True,size=14);cur+=2
        for r in dataframe_to_rows(miss_erp,index=False,header=True): ws.append(r)
        hdr(ws,cur,"AD1457")
    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width=max(len(str(c.value)) if c.value else 0 for c in col)+3
    buf=BytesIO();wb.save(buf);buf.seek(0);return buf

# ==================== MAIN APP ==========================
st.markdown("### Upload Your Files")
erp_up=st.file_uploader("ERP Export (Excel)",type=["xlsx"],key="erp")
ven_up=st.file_uploader("Vendor Statement (Excel)",type=["xlsx"],key="vendor")

if erp_up and ven_up:
    try:
        erp_raw=pd.read_excel(erp_up,dtype=str)
        ven_raw=pd.read_excel(ven_up,dtype=str)
        erp_df=normalize_columns(erp_raw,"erp")
        ven_df=normalize_columns(ven_raw,"ven")

        with st.spinner("Analyzing invoices..."):
            tier1,miss_erp,miss_ven=match_invoices(erp_df,ven_df)
            used_erp=set(tier1["ERP Invoice"].astype(str)) if not tier1.empty else set()
            used_ven=set(tier1["Vendor Invoice"].astype(str)) if not tier1.empty else set()

            tier2,_,_,miss_erp2,miss_ven2=tier2_match(miss_erp,miss_ven)
            if not tier2.empty:
                used_erp|=set(tier2["ERP Invoice"].astype(str))
                used_ven|=set(tier2["Vendor Invoice"].astype(str))
            miss_erp2=miss_erp2[~miss_erp2["Invoice"].astype(str).isin(used_erp)]
            miss_ven2=miss_ven2[~miss_ven2["Invoice"].astype(str).isin(used_ven)]

            tier3,_,_,final_erp_miss,final_ven_miss=tier3_match(miss_erp2,miss_ven2)
            if not tier3.empty:
                used_erp|=set(tier3["ERP Invoice"].astype(str))
                used_ven|=set(tier3["Vendor Invoice"].astype(str))

            # ✅ FINAL CLEANUP
            final_erp_miss=final_erp_miss[~final_erp_miss["Invoice"].astype(str).isin(used_erp)]
            final_ven_miss=final_ven_miss[~final_ven_miss["Invoice"].astype(str).isin(used_ven)]

            erp_pay,ven_pay,pay_match=extract_payments(erp_df,ven_df)

        st.success("Reconciliation Complete!")

        st.markdown('<h2 class="section-title">Missing Invoices</h2>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.subheader("Missing in ERP")
            if not final_ven_miss.empty: st.dataframe(style(final_ven_miss,"background:#AD1457;color:#fff;font-weight:bold;"),use_container_width=True)
            else: st.success("All vendor invoices found.")
        with c2:
            st.subheader("Missing in Vendor")
            if not final_erp_miss.empty: st.dataframe(style(final_erp_miss,"background:#C62828;color:#fff;font-weight:bold;"),use_container_width=True)
            else: st.success("All ERP invoices found.")

        st.markdown('<h2 class="section-title">Download Report</h2>',unsafe_allow_html=True)
        buf=export_excel(final_erp_miss,final_ven_miss)
        st.download_button("Download Excel Report",data=buf,file_name="ReconRaptor_Report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check your Excel columns.")
