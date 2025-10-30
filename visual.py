# ===============================================================
# Overdue Invoices – Priority Vendors Dashboard (FINAL CLICKABLE)
# Features:
# - YES + BD keyword filters
# - AY=0 filter
# - BFP radio mode
# - Clickable Plotly chart (click to show vendor details)
# - Email copy to clipboard
# ===============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import io

# === MUST BE FIRST STREAMLIT COMMAND ===
st.set_page_config(page_title="Overdue Invoices", layout="wide")

st.title("Overdue Invoices – Priority Vendors Dashboard")
st.markdown("**Click a bar segment to view vendor details below.**")

# --- Session State ---
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None
if 'clicked_status' not in st.session_state:
    st.session_state.clicked_status = None

# --- Upload ---
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()

            # Include AY (50) + BT (71)
            keep_cols = [0, 1, 4, 6, 10, 29, 30, 31, 33, 35, 39, 50, 55, 71]
            df_raw = pd.read_excel(
                xls, sheet_name='Outstanding Invoices IB', header=None, usecols=keep_cols
            )

        # --- Find header row ---
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        # --- Assign column names ---
        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Alt_Document', 'Vendor_Email', 'Account_Email',
            'AF', 'AH', 'AJ', 'AN', 'AY', 'BD', 'BT'
        ]

        # --- FILTERS ---
        yes_mask = (
            (df['AF'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AH'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AJ'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AN'].astype(str).str.strip().str.upper() == 'YES')
        )

        bd_keywords = ['ENTERTAINMENT', 'FALSE', 'REGULAR', 'PRIORITY VENDOR', 'PRIORITY VENDOR OS&E']
        bd_mask = df['BD'].astype(str).str.upper().apply(lambda x: any(k in x for k in bd_keywords))

        # AY=0 filter
        def is_zero(v):
            try:
                v = str(v).replace(",", ".").strip()
                return float(v) == 0.0
            except:
                return False

        ay_mask = df['AY'].apply(is_zero)

        df = df[yes_mask & bd_mask & ay_mask].reset_index(drop=True)

        # --- Clean data ---
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No valid invoices after filters.")
            st.stop()

        # --- Overdue logic ---
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # --- RADIO MODE ---
        mode = st.radio("Select Mode:", ["Priority Vendors", "BFP Only"], horizontal=True)
        if mode == "BFP Only":
            df = df[df['BT'].astype(str).str.upper().str.contains("BFP", na=False)]
            if df.empty:
                st.warning("No BFP invoices found.")
                st.stop()

        # --- SUMMARY ---
        summary = (
            df.groupby(['Vendor_Name', 'Status'])['Open_Amount']
            .sum().unstack(fill_value=0).reset_index()
        )
        for col in ['Overdue', 'Not Overdue']:
            if col not in summary.columns:
                summary[col] = 0
        summary['Total'] = summary['Overdue'] + summary['Not Overdue']

        # --- FILTER OPTIONS ---
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show:", ["All Open", "Overdue Only", "Not Overdue Only"])
        with col2:
            top_n_option = st.selectbox("Top N Vendors:", ["Top 20", "Top 30"])

        n = 30 if top_n_option == "Top 30" else 20

        # --- TOP N LOGIC ---
        if status_filter == "All Open":
            top_df = summary.nlargest(n, 'Total').copy()
            title = f"{top_n_option} Vendors (All Open)"
        elif status_filter == "Overdue Only":
            top_df = summary.nlargest(n, 'Overdue').copy()
            top_df['Not Overdue'] = 0
            title = f"{top_n_option} Vendors (Overdue Only)"
        else:
            top_df = summary.nlargest(n, 'Not Overdue').copy()
            top_df['Overdue'] = 0
            title = f"{top_n_option} Vendors (Not Overdue Only)"

        # --- EMAIL EXTRACTION ---
        st.markdown("### 📧 Extract Vendor Emails for Outlook")
        vendor_subset = df[df['Vendor_Name'].isin(top_df['Vendor_Name'])].copy()
        emails = pd.concat([
            vendor_subset['Vendor_Email'],
            vendor_subset['Account_Email']
        ], ignore_index=True).dropna().unique().tolist()
        emails = [e.strip() for e in emails if e.strip() and e.lower() != "nan"]
        email_list = "; ".join(sorted(set(emails)))

        if email_list:
            st.text_area("Emails (ready to copy):", email_list, height=150)
            st.markdown(
                f"""
                <button onclick="navigator.clipboard.writeText('{email_list}')"
                style="background-color:#1f4e79;color:white;border:none;
                padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:bold;">
                📋 Copy to Clipboard</button>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.info("No emails found for this selection.")

        # --- PLOT DATA ---
        plot_df = top_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0].copy()
        plot_df['Status_Label'] = plot_df['Type']

        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title + (" – BFP Only" if mode == "BFP Only" else ""),
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            height=max(500, len(plot_df) * 45),
            custom_data=['Status_Label']
        )

        fig.update_layout(
            xaxis_title="Amount (€)",
            yaxis_title="Vendor",
            legend_title="Status",
            barmode='stack',
            hovermode='y unified',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=160, r=50, t=80, b=50)
        )

        chart = st.plotly_chart(fig, use_container_width=True, key="vendor_chart")

        # --- HANDLE CLICK EVENTS ---
        clicked_vendor = st.session_state.get("clicked_vendor")
        clicked_status = st.session_state.get("clicked_status")

        click_data = chart._get_delta_path() if hasattr(chart, "_get_delta_path") else None
        if hasattr(st, "plotly_chart"):
            # Capture selection dynamically
            if st.session_state.get("vendor_chart") and hasattr(st.session_state["vendor_chart"], "selection"):
                sel = st.session_state["vendor_chart"].selection
                if sel and 'points' in sel and sel['points']:
                    point = sel['points'][0]
                    clicked_vendor = point.get('y')
                    clicked_status = point.get('customdata', [''])[0] if 'customdata' in point else None

        st.session_state.clicked_vendor = clicked_vendor
        st.session_state.clicked_status = clicked_status

        # --- DISPLAY RAW DATA ---
        show_vendor = st.session_state.clicked_vendor
        show_status = st.session_state.clicked_status

        if show_vendor and show_status:
            st.markdown("---")
            st.subheader(f"Raw Data: **{show_vendor}** ({show_status})")

            mask = (df['Vendor_Name'] == show_vendor) & (df['Status'] == show_status)
            raw_details = df[mask].copy()

            if raw_details.empty:
                st.info("No invoices in this segment.")
            else:
                raw_details = raw_details[[
                    'VAT_ID', 'Due_Date', 'Open_Amount', 'Status',
                    'Alt_Document', 'Vendor_Email', 'Account_Email'
                ]]
                raw_details['Due_Date'] = raw_details['Due_Date'].dt.strftime('%Y-%m-%d')
                raw_details['Open_Amount'] = raw_details['Open_Amount'].map('€{:,.2f}'.format)
                st.dataframe(raw_details, use_container_width=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    raw_details.to_excel(writer, index=False, sheet_name='Raw_Data')
                buffer.seek(0)
                st.download_button(
                    f"Download {show_status}",
                    data=buffer,
                    file_name=f"{show_vendor.replace(' ', '_')}_{show_status}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("Click any bar segment to view vendor invoices below.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.stop()

else:
    st.info("Upload Excel to start.")
