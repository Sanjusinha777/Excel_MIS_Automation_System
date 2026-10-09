"""
Interactive Web Application for Excel MIS Report Automation
Built with Streamlit, Pandas, and OpenPyXL
"""

import os
import io
from datetime import datetime
import streamlit as st
import pandas as pd
import altair as alt
from mis_engine import MISEngine
from excel_reporter import ExcelMISReporter
import generate_sample_data

# Page Setup
st.set_page_config(
    page_title="Excel MIS Report Automation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1B365D;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.0rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .kpi-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .kpi-title {
        font-size: 0.8rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1B365D;
    }
    .kpi-sub {
        font-size: 0.75rem;
        color: #94A3B8;
        margin-top: 4px;
    }
    .insight-box {
        background-color: #F0F7FF;
        border-left: 4px solid #2E75B6;
        padding: 14px 18px;
        border-radius: 4px;
        margin-bottom: 20px;
    }
    .stDownloadButton button {
        background-color: #1B365D !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.5rem !important;
        border-radius: 6px !important;
        border: none !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- SIDEBAR -----------------
st.sidebar.image("https://img.icons8.com/color/96/microsoft-excel-2019--v1.png", width=64)
st.sidebar.title("MIS Automation")
st.sidebar.caption("Automated Multi-Tab Management Reporting")

# File Upload / Demo Selector
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel or CSV file",
    type=["xlsx", "xls", "csv"],
    help="Upload any business data file (Sales, Inventory, Finance, HR, Operations, etc.)"
)

# Demo Data Generator Button
demo_file_path = "Sample_Sales_MIS.xlsx"
if st.sidebar.button("⚡ Load Demo Sales MIS Data", use_container_width=True):
    if not os.path.exists(demo_file_path):
        demo_file_path = generate_sample_data.generate_sales_mis_data(500, "Sample_Sales_MIS.xlsx")
    st.session_state["use_demo"] = True
    st.session_state["demo_path"] = demo_file_path

# Choose data source
data_source = None
if uploaded_file is not None:
    data_source = uploaded_file
    st.session_state["use_demo"] = False
elif st.session_state.get("use_demo") and os.path.exists(demo_file_path):
    data_source = demo_file_path

# Main Execution Flow
if data_source is None:
    # Landing / Welcome Screen
    st.markdown("<div class='main-header'>Excel MIS Report Automation</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-header'>Upload any Excel sheet or CSV to instantly generate comprehensive executive MIS analytics and formatted Excel reports.</div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([3, 2])
    with col1:
        st.info("👈 **Get Started:** Upload any **Excel (.xlsx / .xls)** or **CSV** file from the sidebar, or click **'⚡ Load Demo Sales MIS Data'** to test immediately.")
        
        st.markdown("""
        ### Key Capabilities & Automation Features:
        - **Intelligent Auto-Detection**: Automatically identifies Numeric Metrics (Sales, Profit, Cost, Units), Categorical Dimensions (Region, Product, Status), and Timelines/Dates.
        - **Executive KPI Scorecard**: Instantly computes Total Volume, Net/Gross Values, Average Ticket Size, and Peak Records.
        - **Multi-Dimensional Breakdowns**: Generates category-level Contribution % and Pareto (80/20 Rule) cumulative share analysis.
        - **Time-Series & MoM Growth**: Analyzes monthly/quarterly trajectories with Month-over-Month (MoM) growth rates.
        - **Data Quality & Exception Audit**: Detects Missing/Null values, Duplicate records, and Statistical Outliers (IQR Method).
        - **1-Click Executive Excel Export**: Compiles a professionally formatted, 5-sheet corporate Excel workbook (`.xlsx`) with formulas, freeze panes, color hierarchy, and auto-fitted column widths.
        """)
        
    with col2:
        st.markdown("### Preview Features")
        preview_data = pd.DataFrame({
            "Report Tab": ["Executive_Summary", "Category_Analysis", "Trends_Monthly", "Data_Quality_Audit", "Cleaned_Data"],
            "Key Contents": [
                "KPI Scorecard, Business Highlights, Snapshot Table",
                "Dimension breakdown, % Contribution, Pareto cumsum",
                "Month-by-month run-rate, MoM Growth rate %",
                "Data health score, missing values, outlier flags",
                "Cleaned raw records with freeze panes & auto-filter"
            ]
        })
        st.dataframe(preview_data, use_container_width=True, hide_index=True)
        
    st.stop()

# Load Engine
try:
    engine = MISEngine(file_source=data_source)
    if len(engine.sheets) > 1:
        selected_sheet = st.sidebar.selectbox("Select Sheet / Tab", engine.sheets, index=engine.sheets.index(engine.sheet_name))
        if selected_sheet != engine.sheet_name:
            engine.load_data(data_source, sheet_name=selected_sheet)
except Exception as e:
    st.error(f"Error loading dataset: {str(e)}")
    st.stop()

# ----------------- SIDEBAR CONFIGURATION -----------------
st.sidebar.markdown("---")
st.sidebar.subheader("Analytical Mapping")

# Metric Selector
primary_metric = st.sidebar.selectbox(
    "Primary Metric (Value / Amount)",
    options=engine.metric_cols if engine.metric_cols else ["No numeric metric found"],
    index=engine.metric_cols.index(engine.primary_metric) if engine.primary_metric in engine.metric_cols else 0
)

# Dimension Selector
primary_dim = st.sidebar.selectbox(
    "Primary Dimension (Category / Segment)",
    options=engine.dimension_cols if engine.dimension_cols else ["No dimension found"],
    index=engine.dimension_cols.index(engine.primary_dimension) if engine.primary_dimension in engine.dimension_cols else 0
)

# Date Selector
primary_date = None
if engine.date_cols:
    primary_date = st.sidebar.selectbox(
        "Date Column (Timeline)",
        options=engine.date_cols,
        index=engine.date_cols.index(engine.primary_date) if engine.primary_date in engine.date_cols else 0
    )

# Top N selector
top_n = st.sidebar.slider("Top N Ranking Items", min_value=3, max_value=25, value=7)

# ----------------- FILTERS -----------------
st.sidebar.markdown("---")
st.sidebar.subheader("Interactive Filters")

working_df = engine.cleaned_df.copy()

# Date filter
if primary_date and primary_date in working_df.columns:
    valid_dates = pd.to_datetime(working_df[primary_date], errors='coerce').dropna()
    if len(valid_dates) > 0:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        if min_date < max_date:
            date_range = st.sidebar.date_input(
                "Filter Date Range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
                s_dt, e_dt = date_range
                working_df = working_df[
                    (pd.to_datetime(working_df[primary_date]).dt.date >= s_dt) &
                    (pd.to_datetime(working_df[primary_date]).dt.date <= e_dt)
                ]

# Dimension filter
if primary_dim and primary_dim in working_df.columns:
    unique_vals = list(working_df[primary_dim].dropna().unique())
    if len(unique_vals) <= 50:
        selected_cats = st.sidebar.multiselect(
            f"Filter {primary_dim}",
            options=unique_vals,
            default=[]
        )
        if selected_cats:
            working_df = working_df[working_df[primary_dim].isin(selected_cats)]

# Generate Styled Excel Report in Memory for Download
reporter = ExcelMISReporter(engine)
excel_bytes = reporter.save_to_bytes()

# ----------------- MAIN CONTENT -----------------
header_col1, header_col2 = st.columns([3, 1.5])
with header_col1:
    st.markdown("<div class='main-header'>Excel MIS Executive Report</div>", unsafe_allow_html=True)
    source_name = getattr(data_source, "name", "Sample_Sales_MIS.xlsx")
    st.markdown(f"<div class='sub-header'>Analyzed File: <b>{source_name}</b> (Sheet: <i>{engine.sheet_name}</i>) | Active Filtered Rows: <b>{len(working_df):,}</b> of <b>{len(engine.cleaned_df):,}</b></div>", unsafe_allow_html=True)

with header_col2:
    st.download_button(
        label="📥 Download Formatted Excel MIS Report",
        data=excel_bytes,
        file_name=f"MIS_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ----------------- KPI CARDS ROW -----------------
kpis = engine.get_summary_kpis(primary_metric=primary_metric, df_subset=working_df)
tot_val = kpis.get("total_value", 0)
avg_val = kpis.get("avg_value", 0)
max_val = kpis.get("max_value", 0)
cnt_val = kpis.get("total_records", 0)

kpi_c1, kpi_c2, kpi_c3, kpi_c4 = st.columns(4)

with kpi_c1:
    st.markdown(f"""
    <div class='kpi-card'>
        <div class='kpi-title'>Total Transactions</div>
        <div class='kpi-value'>{cnt_val:,}</div>
        <div class='kpi-sub'>Filtered records volume</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_c2:
    st.markdown(f"""
    <div class='kpi-card'>
        <div class='kpi-title'>Total {primary_metric}</div>
        <div class='kpi-value'>{tot_val:,.2f}</div>
        <div class='kpi-sub'>Cumulative metric sum</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_c3:
    st.markdown(f"""
    <div class='kpi-card'>
        <div class='kpi-title'>Average {primary_metric}</div>
        <div class='kpi-value'>{avg_val:,.2f}</div>
        <div class='kpi-sub'>Per transaction mean</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_c4:
    st.markdown(f"""
    <div class='kpi-card'>
        <div class='kpi-title'>Peak {primary_metric}</div>
        <div class='kpi-value'>{max_val:,.2f}</div>
        <div class='kpi-sub'>Maximum recorded value</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# ----------------- EXECUTIVE INSIGHTS CALLOUT -----------------
insights = engine.generate_executive_insights()
insights_html = "".join([f"<li>{ins}</li>" for ins in insights])
st.markdown(f"""
<div class='insight-box'>
    <b style='color: #1B365D; font-size: 1.05rem;'>💡 Automated MIS Executive Insights & Highlights:</b>
    <ul style='margin-top: 8px; margin-bottom: 2px; padding-left: 20px; color: #1E293B;'>
        {insights_html}
    </ul>
</div>
""", unsafe_allow_html=True)

# ----------------- TABS SECTION -----------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Executive Dashboard",
    "📑 Category Breakdown",
    "📈 Time-Series & Trends",
    "🛡️ Data Integrity Audit",
    "📋 Raw Data Explorer"
])

# ---------------- TAB 1: EXECUTIVE DASHBOARD ----------------
with tab1:
    chart_col1, chart_col2 = st.columns(2)
    
    # 1. Primary Dimension Bar Chart
    with chart_col1:
        st.subheader(f"Performance by {primary_dim}")
        dim_summary = engine.get_dimension_breakdown(primary_dim, metric_col=primary_metric, top_n=top_n, df_subset=working_df)
        
        if len(dim_summary) > 0 and 'Total' in dim_summary.columns:
            # Altair Bar Chart
            chart = alt.Chart(dim_summary).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4, color="#1B365D").encode(
                x=alt.X('Total:Q', title=f"Total {primary_metric}"),
                y=alt.Y(f'{primary_dim}:N', sort='-x', title=primary_dim),
                tooltip=[primary_dim, alt.Tooltip('Total:Q', format=",.2f"), alt.Tooltip('Share_Pct:Q', title="Share %", format=".1f")]
            ).properties(height=340)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Insufficient categorical data for dimension chart.")

    # 2. Time Trend or Secondary Dimension Chart
    with chart_col2:
        if primary_date and primary_date in working_df.columns:
            st.subheader("Monthly Run-Rate & Trend")
            trend_df = engine.get_time_series_trend(date_col=primary_date, metric_col=primary_metric, freq="M", df_subset=working_df)
            if len(trend_df) > 0:
                line_chart = alt.Chart(trend_df).mark_line(point=True, color="#2E75B6", strokeWidth=3).encode(
                    x=alt.X('Period:N', title="Month"),
                    y=alt.Y('Total:Q', title=f"Total {primary_metric}"),
                    tooltip=['Period', alt.Tooltip('Total:Q', format=",.2f"), alt.Tooltip('Growth_Pct:Q', title="MoM Growth %", format="+.1f")]
                ).properties(height=340)
                st.altair_chart(line_chart, use_container_width=True)
            else:
                st.info("No timeline records available.")
        elif len(engine.dimension_cols) > 1:
            sec_dim = [d for d in engine.dimension_cols if d != primary_dim][0]
            st.subheader(f"Breakdown by {sec_dim}")
            sec_df = engine.get_dimension_breakdown(sec_dim, metric_col=primary_metric, top_n=top_n, df_subset=working_df)
            if len(sec_df) > 0:
                chart2 = alt.Chart(sec_df).mark_bar(color="#2E75B6").encode(
                    x=alt.X('Total:Q', title=f"Total {primary_metric}"),
                    y=alt.Y(f'{sec_dim}:N', sort='-x', title=sec_dim),
                    tooltip=[sec_dim, alt.Tooltip('Total:Q', format=",.2f"), 'Share_Pct:Q']
                ).properties(height=340)
                st.altair_chart(chart2, use_container_width=True)
        else:
            st.info("No secondary dimension available.")

    # 3. Two-Way Cross Tabulation / Pivot
    st.subheader("Interactive 2-Way Cross-Tabulation (Pivot Table)")
    piv_col1, piv_col2 = st.columns(2)
    with piv_col1:
        dim_row = st.selectbox("Pivot Rows", options=engine.dimension_cols, index=0)
    with piv_col2:
        dim_col_opts = [d for d in engine.dimension_cols if d != dim_row]
        dim_col = st.selectbox("Pivot Columns", options=dim_col_opts if dim_col_opts else engine.dimension_cols, index=0 if dim_col_opts else 0)

    if dim_row and dim_col and dim_row != dim_col:
        pvt = engine.get_cross_tab(dim_row, dim_col, metric_col=primary_metric, df_subset=working_df)
        st.dataframe(pvt.style.format("{:,.2f}").background_gradient(cmap="Blues"), use_container_width=True)

# ---------------- TAB 2: CATEGORY BREAKDOWN ----------------
with tab2:
    st.subheader("Multi-Dimensional Deep-Dive Summaries")
    st.caption("Each category table includes total contribution, average ticket value, transaction volume, and Pareto cumulative %.")
    
    selected_dim_tab = st.selectbox("Select Dimension to Inspect", options=engine.dimension_cols)
    if selected_dim_tab:
        dim_full_df = engine.get_dimension_breakdown(selected_dim_tab, metric_col=primary_metric, top_n=20, df_subset=working_df)
        
        # Display nicely formatted dataframe
        st.dataframe(
            dim_full_df.style.format({
                "Total": "{:,.2f}",
                "Average": "{:,.2f}",
                "Count": "{:,}",
                "Share_Pct": "{:.2f}%",
                "Cumulative_Pct": "{:.2f}%"
            }),
            use_container_width=True
        )

# ---------------- TAB 3: TIME-SERIES & TRENDS ----------------
with tab3:
    if primary_date and primary_date in working_df.columns:
        st.subheader("Monthly & Quarterly Trend Analysis")
        freq_choice = st.radio("Aggregation Frequency", ["Monthly (M)", "Quarterly (Q)", "Yearly (Y)"], horizontal=True)
        freq_code = "M" if "Monthly" in freq_choice else ("Q" if "Quarterly" in freq_choice else "Y")
        
        t_df = engine.get_time_series_trend(date_col=primary_date, metric_col=primary_metric, freq=freq_code, df_subset=working_df)
        
        if len(t_df) > 0:
            st.dataframe(
                t_df.style.format({
                    "Total": "{:,.2f}",
                    "Count": "{:,}",
                    "Avg": "{:,.2f}",
                    "Growth_Pct": "{:+.2f}%"
                }),
                use_container_width=True
            )
            
            # Growth bar chart
            growth_chart = alt.Chart(t_df).mark_bar().encode(
                x=alt.X('Period:N', title="Period"),
                y=alt.Y('Growth_Pct:Q', title="Growth Rate (%)"),
                color=alt.condition(
                    alt.datum.Growth_Pct >= 0,
                    alt.value("#2E7D32"),  # Green positive
                    alt.value("#D32F2F")   # Red negative
                ),
                tooltip=['Period', alt.Tooltip('Growth_Pct:Q', format="+.2f")]
            ).properties(height=260)
            st.altair_chart(growth_chart, use_container_width=True)
    else:
        st.info("No date/time column detected in this dataset.")

# ---------------- TAB 4: DATA QUALITY AUDIT ----------------
with tab4:
    st.subheader("Data Integrity & Exception Audit")
    audit = engine.audit_data_quality()
    
    aud_col1, aud_col2, aud_col3 = st.columns(3)
    aud_col1.metric("Overall Completeness", f"{audit.get('completeness_pct', 100)}%")
    aud_col2.metric("Duplicate Rows", f"{audit.get('duplicate_rows', 0):,}")
    aud_col3.metric("Missing Cell Count", f"{audit.get('missing_cells', 0):,}")
    
    st.markdown("#### Columns with Missing Values")
    miss_df = pd.DataFrame(audit.get("missing_columns", []))
    if len(miss_df) > 0:
        st.dataframe(miss_df, use_container_width=True)
    else:
        st.success("✅ No null or missing values found across all columns.")
        
    st.markdown("#### Statistical Outliers (IQR Method)")
    out_df = pd.DataFrame(audit.get("outliers", []))
    if len(out_df) > 0:
        st.dataframe(out_df, use_container_width=True)
    else:
        st.success("✅ No extreme statistical outliers detected.")

# ---------------- TAB 5: RAW DATA EXPLORER ----------------
with tab5:
    st.subheader("Cleaned Dataset Explorer")
    st.caption(f"Showing first 500 rows of {len(working_df):,} total rows.")
    st.dataframe(working_df.head(500), use_container_width=True)
