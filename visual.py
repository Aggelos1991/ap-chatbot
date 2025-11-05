import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Click a bar → Filter by vendor | Click outside → Reset to all | Table and emails auto-filter**")

# --- SESSION STATE ---
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None

uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx'])

if uploaded_file:
    try:
        # --- LOAD RAW DATA ---
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None)

        # --- HEADER DETECTION ---
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()
        start_row = header_row[0] + 1

        # --- BASE DF (key columns) ---
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)
        df = df.iloc[:, [0, 1, 4, 6, 29, 30, 31, 33, 35, 39]]
        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Vendor_Email', 'Account_Email', 'Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN'
        ]

        # --- BASIC CLEAN ---
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

        # --- ADVANCED FILTERS ---
        st.markdown("### Advanced Filters")

        try:
            headers = df_raw.iloc[header_row[0]].astype(str).str.strip().tolist()
            col_map = {h.upper().strip(): i for i, h in enumerate(headers)}

            bt_idx = next((i for name, i in col_map.items() if "BT" in name and "FUNC" not in name), 42)
            bs_idx = next((i for name, i in col_map.items() if "BS" in name and "FUNC" not in name), 50)
            ba_idx = next((i for name, i in col_map.items() if "BA" in name), 51)

            df['Col_BT'] = df_raw.iloc[start_row:, bt_idx].astype(str).str.strip()
            df['Col_BS'] = df_raw.iloc[start_row:, bs_idx].astype(str).str.strip()
            df['Col_BA'] = df_raw.iloc[start_row:, ba_idx].astype(str).str.strip()

            # --- BT FIX ---
            bt_unique = set(df['Col_BT'].dropna().str.upper().unique())
            if bt_unique <= {"YES", "NO"} or bt_unique <= {"Y", "N"}:
                bt_func_idx = bt_idx + 1
                if bt_func_idx < df_raw.shape[1]:
                    df['Col_BT'] = df_raw.iloc[start_row:, bt_func_idx].astype(str).str.strip()

            # --- BS FIX ---
            def normalize_bs(x):
                x = str(x).strip().upper()
                if x in ["", "OK", "FREE", "0", "FREE FOR PAYMENT"]:
                    return "FREE FOR PAYMENT"
                elif "BLOCK" in x or x in ["1", "BFP"]:
                    return "BFP"
                else:
                    return "FREE FOR PAYMENT"
            df['Col_BS'] = df['Col_BS'].apply(normalize_bs)

        except Exception as e:
            st.warning(f"Couldn't locate BT/BS/BA columns: {e}")
            df['Col_BT'], df['Col_BS'], df['Col_BA'] = "Unknown", "Unknown", "Unknown"

        # --- YES FILTERS ---
        for col in ['Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN']:
            df[col] = df[col].astype(str).str.strip().str.lower()

        apply_yes = st.checkbox("Filter AF/AH/AJ/AN to 'Yes' only", value=True)
        if apply_yes:
            for col in ['Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN']:
                df = df[df[col] == 'yes']

        # --- BT MULTISELECT ---
        bt_values = sorted({v.strip() for v in df['Col_BT'] if v and v.lower() not in ["nan", "none"]})
        selected_bt = st.multiselect("Exceptions / Priority Vendors", bt_values, default=bt_values)
        if selected_bt:
            df = df[df['Col_BT'].isin(selected_bt)]

        # --- BS MULTISELECT ---
        bs_values = sorted({v.strip() for v in df['Col_BS'] if v and v.lower() not in ["nan", "none"]})
        selected_bs = st.multiselect("BFP Status (BS)", bs_values, default=bs_values)
        if selected_bs:
            df = df[df['Col_BS'].isin(selected_bs)]

        # --- COUNTRY FILTER ---
        def classify_country(x):
            x = str(x).strip().lower()
            if "spain" in x or "espa" in x:
                return "Spain"
            return "Foreign"

        df['Country_Type'] = df['Col_BA'].apply(classify_country)
        country_choice = st.radio("Select Country Group", ["All", "Spain", "Foreign"], horizontal=True)
        if country_choice != "All":
            df = df[df['Country_Type'] == country_choice]

        # --- STOP IF NO DATA ---
        if df.empty:
            st.error("No valid vendor data left after filters. Relax your filters or check Excel file.")
            st.stop()

        # --- OVERDUE LOGIC ---
        today = pd.Timestamp.now(tz='Europe/Athens').date()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = np.where(df['Overdue'], 'Overdue', 'Not Overdue')

        # --- SUMMARY ---
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

        # --- FILTERS FOR CHART ---
        c1, c2 = st.columns(2)
        with c1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"])
        with c2:
            vendor_select = st.selectbox("Vendors", ["Top 20", "Top 30"] + sorted(df['Vendor_Name'].unique()))

        top_n = 30 if "30" in vendor_select else 20
        if status_filter == "All Open":
            top_df = summary.nlargest(top_n, 'Total')
        elif status_filter == "Overdue Only":
            top_df = summary.nlargest(top_n, 'Overdue').assign(**{'Not Overdue': 0})
        else:
            top_df = summary.nlargest(top_n, 'Not Overdue').assign(**{'Overdue': 0})
        base_df = top_df if "Top" in vendor_select else summary[summary['Vendor_Name'] == vendor_select]

        # --- CHART ---
        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type', value_name='Amount'
        ).query("Amount>0")

        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            title=f"Top {top_n} Vendors ({status_filter}) — {len(selected_bt)} BT | {len(selected_bs)} BS | {country_choice}",
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
        chart = st.plotly_chart(fig, use_container_width=True, on_select="rerun")

        # --- CLICK HANDLING ---
        if chart.selection and chart.selection['points']:
            st.session_state.clicked_vendor = chart.selection['points'][0].get('y')
        else:
            st.session_state.clicked_vendor = None

        clicked_vendor = st.session_state.clicked_vendor

        # --- FILTER TABLE ---
        filtered_df = df.copy()
        if status_filter == "Overdue Only":
            filtered_df = filtered_df[filtered_df['Status'] == "Overdue"]
        elif status_filter == "Not Overdue Only":
            filtered_df = filtered_df[filtered_df['Status'] == "Not Overdue"]
        if clicked_vendor:
            filtered_df = filtered_df[filtered_df['Vendor_Name'] == clicked_vendor]

        # --- RAW TABLE ---
        if not filtered_df.empty:
            st.subheader("Raw Invoices")
            show = filtered_df[['Vendor_Name','VAT_ID','Due_Date','Open_Amount','Status',
                                'Vendor_Email','Account_Email','Col_AF','Col_AH','Col_AJ','Col_AN',
                                'Col_BT','Col_BS','Col_BA','Country_Type']].copy()
            show['Due_Date'] = pd.to_datetime(show['Due_Date']).dt.strftime("%Y-%m-%d")
            show['Open_Amount'] = show['Open_Amount'].map('€{:,.2f}'.format)
            st.dataframe(show, use_container_width=True)
        else:
            st.info("Click a bar to filter by vendor or adjust filters above.")

        # --- EMAILS ---
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
