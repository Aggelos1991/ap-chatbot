import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from streamlit_plotly_events import plotly_events

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.markdown("""
<h1 style='text-align:center;
    font-family: "Cinzel Decorative", serif;
    font-size: 38px;
    color: goldenrod;
    text-shadow: 2px 2px 6px #000000;'>
💍 One Invoice to rule them all,<br>
One Invoice to seek and find them,<br>
One Invoice to bring them all,<br>
And in the Ledger bind them ⚔️<br>
<span style='font-size:24px; color:#d4af37;'>In the realm of Overdues, where the balances lie. 📜</span>
</h1>

<link href="https://fonts.googleapis.com/css2?family=Cinzel+Decorative:wght@700&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)


# --- SESSION STATE ---
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None

uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx'])

if uploaded_file:
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            # === MAIN SHEET ===
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None)

            # === REFERENCE SHEET (Vendor Type) ===
            if 'VR CHECK_Special vendors list' in xls.sheet_names:
                df_ref = pd.read_excel(xls, sheet_name='VR CHECK_Special vendors list', usecols='A:F', header=None)
                df_ref.columns = ['Vendor_TaxID', 'Col_B', 'Col_C', 'Col_D', 'Col_E', 'Vendor_Category']
            else:
                df_ref = pd.DataFrame(columns=['Vendor_TaxID', 'Vendor_Category'])
                st.warning("Sheet 'VR CHECK_Special vendors list' not found. Vendor categories may be missing.")

            # === COUNTRY LOOKUP SHEET ===
            if 'Vendors' in xls.sheet_names:
                df_country = pd.read_excel(xls, sheet_name='Vendors', usecols='A:G', header=None)
                df_country.columns = ['Vendor_TaxID', 'Col_B', 'Col_C', 'Col_D', 'Col_E', 'Col_F', 'Country']
                df_country['Vendor_TaxID'] = df_country['Vendor_TaxID'].astype(str).str.strip().str.upper()
            else:
                df_country = pd.DataFrame(columns=['Vendor_TaxID', 'Country'])
                st.warning("Sheet 'Vendors' not found. Country classification may be incomplete.")

        # === HEADER DETECTION ===
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()
        start_row = header_row[0] + 1

        # === BASE COLUMNS ===
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)
        df = df.iloc[:, [0, 1, 4, 6, 29, 30, 31, 33, 35, 39]]
        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Vendor_Email', 'Account_Email', 'Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN'
        ]

        # === CLEAN ===
        df = df.dropna(how="all")
        df = df[df['Vendor_Name'].notna()]
        df = df[~df['Vendor_Name'].astype(str).str.strip().eq("")]
        bad_patterns = r"(?i)total|saldo|asiento|header|proveedor|unnamed|vendor|facturas|periodo|sum|importe"
        df = df[~df['Vendor_Name'].astype(str).str.contains(bad_patterns, na=False)]
        df = df[~df['Open_Amount'].astype(str).str.contains(bad_patterns, na=False)]
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce').dt.date
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Due_Date', 'Open_Amount'])
        df = df[df['Open_Amount'] > 0]

        # === FIND BS / BA COLUMNS ===
        headers = df_raw.iloc[header_row[0]].astype(str).str.strip().tolist()
        col_map = {h.upper().strip(): i for i, h in enumerate(headers)}
        bs_idx = next((i for name, i in col_map.items() if "BS" in name and "FUNC" not in name), 50)
        ba_idx = next((i for name, i in col_map.items() if "BA" in name), 51)
        df['Col_BS'] = df_raw.iloc[start_row:, bs_idx].astype(str).str.strip()
        df['Col_BA'] = df_raw.iloc[start_row:, ba_idx].astype(str).str.strip()

        # === MERGE VENDOR TYPE ===
        if not df_ref.empty:
            df_ref['Vendor_TaxID'] = df_ref['Vendor_TaxID'].astype(str).str.strip().str.upper()
            df['VAT_ID_clean'] = df['VAT_ID'].astype(str).str.strip().str.upper()
            df = df.merge(df_ref[['Vendor_TaxID', 'Vendor_Category']],
                          left_on='VAT_ID_clean', right_on='Vendor_TaxID', how='left')
            df['Vendor_Type'] = df['Vendor_Category'].fillna("Uncategorized")
        else:
            df['Vendor_Type'] = "Uncategorized"

        # === MERGE COUNTRY INFO ===
        df['VAT_ID_clean'] = df['VAT_ID'].astype(str).str.strip().str.upper()
        df = df.merge(df_country[['Vendor_TaxID', 'Country']],
                      left_on='VAT_ID_clean', right_on='Vendor_TaxID', how='left')
        df['Country_Type'] = df['Country'].fillna("Unknown").apply(
            lambda x: "Spain" if isinstance(x, str) and "spain" in x.lower()
            else "Foreign" if isinstance(x, str) and x.strip() != ""
            else "Unknown"
        )

        # === NORMALIZE BFP ===
        def normalize_bs(x):
            x = str(x).strip().upper()
            if x in ["", "OK", "FREE", "0", "FREE FOR PAYMENT"]:
                return "Free for Payment"
            elif "BLOCK" in x or x in ["1", "BFP"]:
                return "Blocked for Payment"
            else:
                return x
        df['Col_BS'] = df['Col_BS'].apply(normalize_bs)

        # === NORMALIZE YES FILTERS ===
        for col in ['Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN']:
            df[col] = df[col].fillna("").astype(str).apply(lambda x: x.strip().lower() if isinstance(x, str) else "")

        # === ADVANCED FILTERS ===
        st.markdown("### Advanced Filters")
        apply_yes = st.checkbox("Filter AF/AH/AN to 'Yes' only", value=True)
        if apply_yes:
            for col in ['Col_AF', 'Col_AH', 'Col_AN']:
                df = df[df[col] == 'yes']

        aj_yes_only = st.checkbox("Filter AJ = 'Yes' only", value=False)
        if aj_yes_only:
            df = df[df['Col_AJ'] == 'yes']

        # === BT FILTER (Vendor_Type) ===
        bt_values = sorted({v.strip() for v in df['Vendor_Type'] if v and v.lower() not in ["nan", "none"]})
        selected_bt = st.multiselect("Exceptions / Priority Vendors (Vendor Type)", bt_values, default=bt_values)
        if selected_bt:
            df = df[df['Vendor_Type'].isin(selected_bt)]

        # === BS FILTER (BFP Status) ===
        bs_values = sorted({v.strip() for v in df['Col_BS'] if v and v.lower() not in ["nan", "none"]})
        selected_bs = st.multiselect("BFP Status (BS)", bs_values, default=bs_values)
        if selected_bs:
            df = df[df['Col_BS'].isin(selected_bs)]

        # === COUNTRY FILTER ===
        country_choice = st.radio("Select Country Group", ["All", "Spain", "Foreign"], horizontal=True)
        if country_choice != "All":
            df = df[df['Country_Type'] == country_choice]

        # === OVERDUE LOGIC ===
        today = pd.Timestamp.now(tz='Europe/Athens').date()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = np.where(df['Overdue'], 'Overdue', 'Not Overdue')

        # === SUMMARY ===
        summary = (
            df.groupby(['Vendor_Name', 'Status'], as_index=False)['Open_Amount']
              .sum()
              .pivot(index='Vendor_Name', columns='Status', values='Open_Amount')
              .fillna(0)
              .reset_index()
        )
        for col in ['Overdue', 'Not Overdue']:
            if col not in summary.columns:
                summary[col] = 0
        summary['Total'] = summary['Overdue'] + summary['Not Overdue']

        # === CHART FILTERS ===
        c1, c2 = st.columns(2)
        with c1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"])
        with c2:
            vendor_select = st.selectbox("Vendors", ["All Vendors", "Top 20", "Top 30"] + sorted(df['Vendor_Name'].unique()))

        top_n = 30 if "30" in vendor_select else 20
        if vendor_select == "All Vendors":
            base_df = summary
        elif "Top" in vendor_select:
            if status_filter == "All Open":
                base_df = summary.nlargest(top_n, 'Total')
            elif status_filter == "Overdue Only":
                base_df = summary.nlargest(top_n, 'Overdue').assign(**{'Not Overdue': 0})
            else:
                base_df = summary.nlargest(top_n, 'Not Overdue').assign(**{'Overdue': 0})
        else:
            base_df = summary[summary['Vendor_Name'] == vendor_select]

        # === CHART ===
        plot_df = base_df.melt(id_vars='Vendor_Name',
                               value_vars=['Overdue', 'Not Overdue'],
                               var_name='Type', value_name='Amount').query("Amount>0")
        fig = px.bar(plot_df, x='Amount', y='Vendor_Name', color='Type',
                     orientation='h', color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
                     title=f"{vendor_select} ({status_filter}) — {country_choice}",
                     height=max(500, len(plot_df) * 45))

        totals = plot_df.groupby('Vendor_Name')['Amount'].sum().reset_index()
        fig.add_scatter(x=totals['Amount'], y=totals['Vendor_Name'], mode='text',
                        text=totals['Amount'].apply(lambda x: f"€{x:,.0f}"),
                        textposition='top center', textfont=dict(size=14, color='white', family='Arial Black'),
                        showlegend=False, hoverinfo='skip')
        fig.update_layout(xaxis_title="Amount (€)", yaxis_title="Vendor", legend_title="Status",
                          barmode='stack', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                          margin=dict(l=150, r=50, t=80, b=50))

        # === CLICK HANDLING (UPDATED) ===
        selected_points = plotly_events(
            fig,
            click_event=True,
            hover_event=False,
            select_event=False,
            override_height=max(500, len(plot_df) * 45),
            key="vendor_chart"
        )

        if selected_points:
            st.session_state.clicked_vendor = selected_points[0]["y"]
        else:
            st.session_state.clicked_vendor = None

        clicked_vendor = st.session_state.clicked_vendor

        # === FILTERED TABLE ===
        filtered_df = df.copy()
        if status_filter == "Overdue Only":
            filtered_df = filtered_df[filtered_df['Status'] == "Overdue"]
        elif status_filter == "Not Overdue Only":
            filtered_df = filtered_df[filtered_df['Status'] == "Not Overdue"]
        if clicked_vendor:
            filtered_df = filtered_df[filtered_df['Vendor_Name'] == clicked_vendor]

        # === TABLE DISPLAY ===
        if not filtered_df.empty:
            st.subheader("Raw Invoices")
            show = filtered_df[['Vendor_Name','VAT_ID','Due_Date','Open_Amount','Status',
                                'Vendor_Email','Account_Email','Col_AF','Col_AH','Col_AJ','Col_AN',
                                'Vendor_Type','Col_BS','Country_Type']].copy()
            show['Due_Date'] = pd.to_datetime(show['Due_Date']).dt.strftime("%Y-%m-%d")
            show['Open_Amount'] = show['Open_Amount'].map('€{:,.2f}'.format)
            st.dataframe(show, use_container_width=True)
        else:
            st.info("Click a bar to filter by vendor or adjust filters above.")

        # === EMAILS ===
        st.markdown("---")
        st.subheader("📧 Emails (copy for Outlook)")
        emails = pd.concat([filtered_df['Vendor_Email'], filtered_df['Account_Email']], ignore_index=True)
        emails = emails.dropna().astype(str)
        emails = emails[emails.str.contains('@')]
        lang = "Spanish" if country_choice == "Spain" else "English" if country_choice == "Foreign" else "Mixed"
        unique = sorted(set(emails))
        st.text_area(f"Ctrl + C to copy ({lang} vendors):", ", ".join(unique), height=120)
        st.success(f"{len(unique)} {lang} emails collected")

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("Upload Excel → Click a bar → Filter data | Click outside → Reset to all")
