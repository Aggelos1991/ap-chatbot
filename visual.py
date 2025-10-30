# ===============================================================
# Overdue Invoices – Priority Vendors Dashboard (CLICKABLE + EMAIL + BFP + AY=0 + Individual Vendor)
# ===============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import io

# === MUST BE FIRST STREAMLIT COMMAND ===
st.set_page_config(page_title="Overdue Invoices", layout="wide")

st.title("Overdue Invoices – Priority Vendors Dashboard")
st.markdown("**Click a bar segment to see only that data | Export to Filtered Excel**")

# Session state
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None
if 'clicked_status' not in st.session_state:
    st.session_state.clicked_status = None

# Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()

            # === Added AY(50) and BT(71) ===
            keep_cols = [0, 1, 4, 6, 10, 29, 30, 31, 33, 35, 39, 50, 55, 71]
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None, usecols=keep_cols)

        # Find header row
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        # Assign column names
        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Alt_Document', 'Vendor_Email', 'Account_Email',
            'AF', 'AH', 'AJ', 'AN', 'AY', 'BD', 'BT'
        ]

        # === Original YES + BD Filter ===
        yes_mask = (
            (df['AF'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AH'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AJ'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AN'].astype(str).str.strip().str.upper() == 'YES')
        )

        bd_keywords = ['ENTERTAINMENT', 'FALSE', 'REGULAR', 'PRIORITY VENDOR', 'PRIORITY VENDOR OS&E']
        bd_mask = df['BD'].astype(str).str.upper().apply(lambda x: any(k in x for k in bd_keywords))

        # === AY = 0 filter ===
        def is_zero(v):
            try:
                v = str(v).replace(",", ".").strip()
                return float(v) == 0.0
            except:
                return False

        ay_mask = df['AY'].apply(is_zero)

        # === Apply combined filters ===
        df = df[yes_mask & bd_mask & ay_mask].reset_index(drop=True)

        # === BFP filter radio ===
        bfp_mode = st.radio("Vendor Mode:", ["All Vendors", "BFP Only"], horizontal=True)
        if bfp_mode == "BFP Only":
            df = df[df['BT'].astype(str).str.upper().str.contains("BFP", na=False)]
            if df.empty:
                st.warning("No BFP invoices found.")
                st.stop()

        df = df.drop(columns=['AF', 'AH', 'AJ', 'AN', 'AY', 'BD', 'BT'])

        # === Clean ===
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No valid open invoices.")
            st.stop()

        # Overdue logic
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # === Summary ===
        summary = (
            df.groupby(['Vendor_Name', 'Status'])['Open_Amount']
            .sum()
            .unstack(fill_value=0)
            .reset_index()
        )
        for col in ['Overdue', 'Not Overdue']:
            if col not in summary.columns:
                summary[col] = 0
        summary['Total'] = summary['Overdue'] + summary['Not Overdue']
        full_summary = summary

        # === Filters ===
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        with col2:
            top_n_option = st.selectbox("Top N", ["Top 20", "Top 30"], key="top_n")

        n = 30 if top_n_option == "Top 30" else 20

        # === Individual Vendor Selector ===
        vendor_list = sorted(df['Vendor_Name'].dropna().unique().tolist())
        selected_vendor = st.selectbox("Or Select Individual Vendor", [""] + vendor_list, key="vendor_select")

        # === TOP N logic ===
        if selected_vendor and selected_vendor != "":
            top_df = full_summary[full_summary['Vendor_Name'] == selected_vendor].copy()
            title = f"Vendor: {selected_vendor}"
        elif status_filter == "All Open":
            top_df = full_summary.nlargest(n, 'Total').copy()
            title = f"{top_n_option} Vendors (All Open)"
        elif status_filter == "Overdue Only":
            top_df = full_summary.nlargest(n, 'Overdue').copy()
            top_df['Not Overdue'] = 0
            title = f"{top_n_option} Vendors (Overdue Only)"
        else:
            top_df = full_summary.nlargest(n, 'Not Overdue').copy()
            top_df['Overdue'] = 0
            title = f"{top_n_option} Vendors (Not Overdue Only)"

        # === EMAIL extraction ===
        st.markdown("### 📧 Extract Vendor Emails for Outlook")
        vendor_subset = df[df['Vendor_Name'].isin(top_df['Vendor_Name'])].copy()
        emails = pd.concat([
            vendor_subset['Vendor_Email'],
            vendor_subset['Account_Email']
        ], ignore_index=True).dropna().unique().tolist()
        emails = [e.strip() for e in emails if e.strip() and e.lower() != "nan"]
        email_list = "; ".join(sorted(set(emails)))

        if email_list:
            st.text_area("Emails (ready to copy):", email_list, height=120)
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

        # === PLOT DATA ===
        plot_df = top_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0].copy()
        plot_df['Status_Label'] = plot_df['Type']

        # === BAR CHART ===
        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title,
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            height=max(500, len(plot_df) * 45),
            custom_data=['Status_Label']
        )

        totals = plot_df.groupby('Vendor_Name')['Amount'].sum().reset_index()
        fig.add_scatter(
            x=totals['Amount'],
            y=totals['Vendor_Name'],
            mode='text',
            text=totals['Amount'].apply(lambda x: f'€{x:,.0f}'),
            textposition='top center',
            textfont=dict(size=14, color='white', family='Arial Black'),
            showlegend=False,
            hoverinfo='skip'
        )

        fig.update_layout(
            xaxis_title="Amount (€)",
            yaxis_title="Vendor",
            legend_title="Status",
            barmode='stack',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=160, r=50, t=80, b=50)
        )

        # === INTERACTIVE CHART ===
        chart = st.plotly_chart(fig, use_container_width=True, key="vendor_chart", on_select="rerun")

        # === CAPTURE CLICK ===
        clicked_vendor = None
        clicked_status = None
        if chart.selection and 'points' in chart.selection and chart.selection['points']:
            point = chart.selection['points'][0]
            if 'y' in point and 'customdata' in point and point['customdata']:
                clicked_vendor = point['y']
                clicked_status = point['customdata'][0]
            elif 'y' in point:
                clicked_vendor = point['y']
                clicked_status = 'Overdue' if point.get('marker.color') == '#8B0000' else 'Not Overdue'

        st.session_state.clicked_vendor = clicked_vendor
        st.session_state.clicked_status = clicked_status

        # === SHOW RAW DATA (CLICKED SEGMENT) ===
        show_vendor = st.session_state.clicked_vendor
        show_status = st.session_state.clicked_status

        if show_vendor and show_status:
            st.markdown("---")
            st.subheader(f"Raw Data: **{show_vendor}** → **{show_status}**")

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
            st.info("**Click any colored bar segment** to view only that data.**")

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.stop()

else:
    st.info("Upload Excel to Click bar segment to See only that data to Export")
