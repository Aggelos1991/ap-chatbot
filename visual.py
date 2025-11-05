import streamlit as st
import pandas as pd
import plotly.express as px
import io
import numpy as np

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Click a bar segment → See only that data | Export → Raw data**")

if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None
if 'clicked_type' not in st.session_state:
    st.session_state.clicked_type = None

uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx'])

if uploaded_file:
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None)

        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)
        df = df.iloc[:, [0, 1, 4, 6, 29, 30]]
        df.columns = ['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']

        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce').dt.date
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        today = pd.Timestamp.now(tz='Europe/Athens').date()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = np.where(df['Overdue'], 'Overdue', 'Not Overdue')

        full_summary = (
            df.groupby(['Vendor_Name', 'Status'], as_index=False)['Open_Amount']
              .sum()
              .pivot(index='Vendor_Name', columns='Status', values='Open_Amount')
              .fillna(0)
              .reset_index()
        )
        for col in ['Overdue', 'Not Overdue']:
            if col not in full_summary.columns:
                full_summary[col] = 0
        full_summary['Total'] = full_summary['Overdue'] + full_summary['Not Overdue']

        c1, c2, c3 = st.columns(3)
        with c1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"])
        with c2:
            vendor_select = st.selectbox("Vendors", ["Top 20", "Top 30"] + sorted(df['Vendor_Name'].unique()))
     

        top_n = 30 if "30" in vendor_select else 20
        if status_filter == "All Open":
            top_df = full_summary.nlargest(top_n, 'Total')
        elif status_filter == "Overdue Only":
            top_df = full_summary.nlargest(top_n, 'Overdue').assign(**{'Not Overdue': 0})
        else:
            top_df = full_summary.nlargest(top_n, 'Not Overdue').assign(**{'Overdue': 0})

        base_df = top_df if "Top" in vendor_select else full_summary[full_summary['Vendor_Name'] == vendor_select]

        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type', value_name='Amount'
        ).query("Amount>0")

        fig = px.bar(
            plot_df, x='Amount', y='Vendor_Name', color='Type', orientation='h',
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            title=f"Top {top_n} Vendors ({status_filter})",
            height=max(500, len(plot_df) * 45)
        )
        totals = plot_df.groupby('Vendor_Name')['Amount'].sum().reset_index()
        fig.add_scatter(
            x=totals['Amount'], y=totals['Vendor_Name'],
            mode='text', text=totals['Amount'].apply(lambda x: f"€{x:,.0f}"),
            textposition='top center',
            textfont=dict(size=14, color='white', family='Arial Black'),
            showlegend=False, hoverinfo='skip'
        )
        fig.update_layout(
            xaxis_title="Amount (€)", yaxis_title="Vendor", legend_title="Status",
            barmode='stack', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=150, r=50, t=80, b=50)
        )

        chart = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key="chart")

        if chart.selection and chart.selection['points']:
            point = chart.selection['points'][0]
            st.session_state.clicked_vendor = point['y']
            st.session_state.clicked_type = point['curveNumber']  # 0 = Overdue, 1 = Not Overdue

        clicked_vendor = st.session_state.clicked_vendor
        clicked_type = st.session_state.clicked_type

        if clicked_vendor:
            clicked_status = "Overdue" if clicked_type == 0 else "Not Overdue"
            st.subheader(f"Raw Invoices – {clicked_vendor} ({clicked_status})")
            raw = df[(df['Vendor_Name'] == clicked_vendor) & (df['Status'] == clicked_status)].copy()
            raw = raw[['VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Vendor_Email', 'Account_Email']]
            raw['Due_Date'] = pd.to_datetime(raw['Due_Date']).dt.strftime("%Y-%m-%d")
            raw['Open_Amount'] = raw['Open_Amount'].map('€{:,.2f}'.format)
            st.dataframe(raw, use_container_width=True)
        else:
            st.info("Click a bar segment to see vendor-specific data.")

        st.markdown("---")
        st.subheader("📧 Emails (copy for Outlook)")
        scope = df if status_filter == "All Open" else df[df['Overdue']] if status_filter == "Overdue Only" else df[~df['Overdue']]
        emails = pd.concat([scope['Vendor_Email'], scope['Account_Email']], ignore_index=True).dropna().astype(str)
        emails = emails[emails.str.contains('@')]
        unique_emails = sorted(set(emails))
        st.text_area("Ctrl + C to copy:", ", ".join(unique_emails), height=120)
        st.success(f"{len(unique_emails)} emails collected")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload Excel → Click bar segment → See data → Copy emails")
