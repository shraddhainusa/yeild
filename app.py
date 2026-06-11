"""
Yield Performance Dashboard
A comprehensive Streamlit dashboard for analyzing yield performance metrics.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')

# Configure page
st.set_page_config(
    page_title="Yield Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
        .metric-card {
            background-color: #f0f2f6;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .kpi-value {
            font-size: 32px;
            font-weight: bold;
            color: #1f77b4;
        }
        .kpi-label {
            font-size: 14px;
            color: #666;
            margin-top: 5px;
        }
    </style>
""", unsafe_allow_html=True)

# ==================== DATA LOADING ====================
@st.cache_data
def load_excel_data(file_path):
    """Load all sheets from the Excel workbook"""
    excel_file = pd.ExcelFile(file_path)
    sheets = {}
    
    for sheet_name in excel_file.sheet_names:
        sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
    
    return sheets

# Load data
data_file = "dummy_yield_history_data_updated.xlsx"
if not Path(data_file).exists():
    st.error(f"❌ File not found: {data_file}")
    st.info("Please upload the Excel file to the same directory as this script.")
    st.stop()

try:
    sheets = load_excel_data(data_file)
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()

# ==================== DATA PREPARATION ====================

# Get individual sheets
df_overall = sheets.get("Yield History Overall", pd.DataFrame())
df_monthly = sheets.get("Yield History Monthly", pd.DataFrame())
df_90_rolling = sheets.get("90 Days Rolling History", pd.DataFrame())
df_90_pivot = sheets.get("90 Day Pivot", pd.DataFrame())

# Reshape Monthly data to long format
def reshape_monthly_data(df):
    """Convert monthly sheet from wide to long format"""
    if df.empty:
        return df
    
    df_copy = df.copy()
    
    # Identify the Material column (first column)
    material_col = df_copy.columns[0]
    df_copy = df_copy.rename(columns={material_col: 'Material'})
    
    # Melt the dataframe
    id_vars = ['Material']
    value_vars = [col for col in df_copy.columns if col != 'Material']
    
    df_melted = df_copy.melt(id_vars=id_vars, value_vars=value_vars, 
                              var_name='Period', value_name='Value')
    
    # Extract metric type from column name
    df_melted['Metric'] = df_melted['Period'].str.extract(r'(MRP Plan|Actual|Variance)')
    df_melted['Period_Date'] = df_melted['Period'].str.replace(r'(MRP Plan|Actual|Variance)', '', regex=True).str.strip()
    
    # Pivot to get separate columns for each metric
    df_reshaped = df_melted.pivot_table(
        index=['Material', 'Period_Date'],
        columns='Metric',
        values='Value',
        aggfunc='first'
    ).reset_index()
    
    return df_reshaped

df_monthly_long = reshape_monthly_data(df_monthly)

# ==================== SIDEBAR FILTERS ====================

st.sidebar.title("🔧 Dashboard Filters")

# Section 1: Executive Summary
with st.sidebar.expander("📊 Executive Summary", expanded=True):
    st.write("**No filters needed**")

# Section 2: Yield History Overall
with st.sidebar.expander("📈 Yield History Overall", expanded=False):
    if not df_overall.empty:
        materials_overall = df_overall[df_overall.columns[0]].unique()
        selected_material_overall = st.multiselect(
            "Select Materials",
            options=materials_overall,
            default=list(materials_overall)[:5] if len(materials_overall) > 0 else [],
            key="materials_overall"
        )
        
        variance_range = st.slider(
            "Variance Range (%)",
            min_value=-100.0,
            max_value=100.0,
            value=(-100.0, 100.0),
            key="variance_range"
        )

# Section 3: Monthly Yield Trend
with st.sidebar.expander("📅 Monthly Yield Trend", expanded=False):
    if not df_monthly_long.empty:
        materials_monthly = df_monthly_long['Material'].unique()
        selected_material_monthly = st.multiselect(
            "Select Materials",
            options=materials_monthly,
            default=list(materials_monthly)[:3] if len(materials_monthly) > 0 else [],
            key="materials_monthly"
        )

# Section 4: 90 Days Rolling History
with st.sidebar.expander("📊 90 Days Rolling History", expanded=False):
    if not df_90_rolling.empty:
        # Try to identify Material column
        material_col_rolling = None
        for col in df_90_rolling.columns:
            if 'material' in str(col).lower() or col == df_90_rolling.columns[0]:
                material_col_rolling = col
                break
        
        if material_col_rolling and material_col_rolling in df_90_rolling.columns:
            materials_rolling = df_90_rolling[material_col_rolling].unique()
            selected_material_rolling = st.multiselect(
                "Select Materials",
                options=materials_rolling,
                default=list(materials_rolling)[:3] if len(materials_rolling) > 0 else [],
                key="materials_rolling"
            )
        
        # Check for Fiscal Year Week column
        year_week_col = None
        for col in df_90_rolling.columns:
            if 'week' in str(col).lower() or 'fiscal' in str(col).lower():
                year_week_col = col
                break
        
        if year_week_col and year_week_col in df_90_rolling.columns:
            year_weeks = sorted(df_90_rolling[year_week_col].unique())
            selected_year_week = st.multiselect(
                "Select Fiscal Year Weeks",
                options=year_weeks,
                default=year_weeks[-4:] if len(year_weeks) > 4 else year_weeks,
                key="year_week"
            )

# ==================== MAIN DASHBOARD ====================

# Title
st.title("📊 Yield Performance Dashboard")
st.markdown("---")

# ==================== SECTION 1: EXECUTIVE SUMMARY ====================

st.header("1️⃣ Executive Summary")

if not df_overall.empty:
    try:
        # Identify column names
        col_names = df_overall.columns.tolist()
        
        # Calculate metrics (try common column name patterns)
        metrics = {}
        
        # Average MRP Plan %
        for col in col_names:
            if 'mrp' in str(col).lower() and 'plan' in str(col).lower():
                metrics['avg_mrp_plan'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
                break
        
        # Average Actual %
        for col in col_names:
            if 'actual' in str(col).lower() and '%' in str(col):
                metrics['avg_actual'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
                break
        
        # Average Variance %
        for col in col_names:
            if 'variance' in str(col).lower():
                metrics['avg_variance'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
                break
        
        # Total Orders
        for col in col_names:
            if 'order' in str(col).lower() and 'qty' not in str(col).lower():
                metrics['total_orders'] = pd.to_numeric(df_overall[col], errors='coerce').sum()
                break
        
        # Total Start Quantity
        for col in col_names:
            if 'start' in str(col).lower() and 'qty' in str(col).lower():
                metrics['start_qty'] = pd.to_numeric(df_overall[col], errors='coerce').sum()
                break
        
        # Total Delivered
        for col in col_names:
            if 'deliver' in str(col).lower() and 'qty' in str(col).lower():
                metrics['delivered_qty'] = pd.to_numeric(df_overall[col], errors='coerce').sum()
                break
        
        # Average Weighted Average %
        for col in col_names:
            if 'weighted' in str(col).lower():
                metrics['avg_weighted'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
                break
        
        # Display metrics in columns
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "📊 Avg MRP Plan %",
                f"{metrics.get('avg_mrp_plan', 0):.2f}%"
            )
        
        with col2:
            st.metric(
                "✅ Avg Actual %",
                f"{metrics.get('avg_actual', 0):.2f}%"
            )
        
        with col3:
            st.metric(
                "📈 Avg Variance %",
                f"{metrics.get('avg_variance', 0):.2f}%"
            )
        
        with col4:
            st.metric(
                "⚖️ Avg Weighted Avg %",
                f"{metrics.get('avg_weighted', 0):.2f}%"
            )
        
        col5, col6, col7 = st.columns(3)
        
        with col5:
            st.metric(
                "📦 Total Orders",
                f"{int(metrics.get('total_orders', 0)):,}"
            )
        
        with col6:
            st.metric(
                "📤 Start Quantity",
                f"{int(metrics.get('start_qty', 0)):,}"
            )
        
        with col7:
            st.metric(
                "📬 Delivered Quantity",
                f"{int(metrics.get('delivered_qty', 0)):,}"
            )
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error calculating metrics: {e}")

# ==================== SECTION 2: YIELD HISTORY OVERALL ====================

st.header("2️⃣ Yield History Overall")

if not df_overall.empty:
    try:
        # Apply filters
        filtered_df = df_overall.copy()
        
        if 'selected_material_overall' in st.session_state and len(st.session_state.selected_material_overall) > 0:
            material_col = df_overall.columns[0]
            filtered_df = filtered_df[filtered_df[material_col].isin(st.session_state.selected_material_overall)]
        
        # Display table
        st.subheader("Overall Yield Data")
        st.dataframe(filtered_df, use_container_width=True)
        
        # Create comparison chart
        st.subheader("MRP Plan vs Actual by Material")
        
        # Find MRP Plan and Actual columns
        mrp_col = None
        actual_col = None
        material_col = df_overall.columns[0]
        
        for col in df_overall.columns:
            if 'mrp' in str(col).lower() and 'plan' in str(col).lower() and not mrp_col:
                mrp_col = col
            if 'actual' in str(col).lower() and '%' in str(col) and not actual_col:
                actual_col = col
        
        if mrp_col and actual_col:
            chart_data = filtered_df[[material_col, mrp_col, actual_col]].copy()
            chart_data[mrp_col] = pd.to_numeric(chart_data[mrp_col], errors='coerce')
            chart_data[actual_col] = pd.to_numeric(chart_data[actual_col], errors='coerce')
            
            fig = go.Figure(data=[
                go.Bar(name='MRP Plan %', x=chart_data[material_col], y=chart_data[mrp_col]),
                go.Bar(name='Actual %', x=chart_data[material_col], y=chart_data[actual_col])
            ])
            
            fig.update_layout(
                title="MRP Plan vs Actual Yield %",
                xaxis_title="Material",
                yaxis_title="Yield %",
                barmode='group',
                hovermode='x unified',
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in Yield History Overall: {e}")

# ==================== SECTION 3: MONTHLY YIELD TREND ====================

st.header("3️⃣ Monthly Yield Trend")

if not df_monthly_long.empty:
    try:
        # Apply filters
        filtered_monthly = df_monthly_long.copy()
        
        if 'selected_material_monthly' in st.session_state and len(st.session_state.selected_material_monthly) > 0:
            filtered_monthly = filtered_monthly[filtered_monthly['Material'].isin(st.session_state.selected_material_monthly)]
        
        # Line chart for MRP Plan vs Actual
        st.subheader("MRP Plan vs Actual Over Time")
        
        if 'MRP Plan' in filtered_monthly.columns and 'Actual' in filtered_monthly.columns:
            fig_trend = go.Figure()
            
            for material in filtered_monthly['Material'].unique():
                mat_data = filtered_monthly[filtered_monthly['Material'] == material]
                
                fig_trend.add_trace(go.Scatter(
                    x=mat_data['Period_Date'],
                    y=mat_data['MRP Plan'],
                    mode='lines+markers',
                    name=f'{material} - MRP Plan'
                ))
                
                fig_trend.add_trace(go.Scatter(
                    x=mat_data['Period_Date'],
                    y=mat_data['Actual'],
                    mode='lines+markers',
                    name=f'{material} - Actual',
                    line=dict(dash='dash')
                ))
            
            fig_trend.update_layout(
                title="MRP Plan vs Actual Yield Over Time",
                xaxis_title="Period",
                yaxis_title="Yield %",
                hovermode='x unified',
                height=400
            )
            
            st.plotly_chart(fig_trend, use_container_width=True)
        
        # Variance chart
        st.subheader("Variance Over Time")
        
        if 'Variance' in filtered_monthly.columns:
            fig_var = go.Figure()
            
            for material in filtered_monthly['Material'].unique():
                mat_data = filtered_monthly[filtered_monthly['Material'] == material]
                
                fig_var.add_trace(go.Bar(
                    x=mat_data['Period_Date'],
                    y=mat_data['Variance'],
                    name=material
                ))
            
            fig_var.update_layout(
                title="Variance Over Time",
                xaxis_title="Period",
                yaxis_title="Variance %",
                hovermode='x unified',
                height=400
            )
            
            st.plotly_chart(fig_var, use_container_width=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in Monthly Yield Trend: {e}")

# ==================== SECTION 4: 90 DAYS ROLLING HISTORY ====================

st.header("4️⃣ 90 Days Rolling History")

if not df_90_rolling.empty:
    try:
        # Apply filters
        filtered_rolling = df_90_rolling.copy()
        
        # Try to find and filter by material
        material_col_rolling = None
        for col in df_90_rolling.columns:
            if 'material' in str(col).lower() or col == df_90_rolling.columns[0]:
                material_col_rolling = col
                break
        
        if material_col_rolling and 'selected_material_rolling' in st.session_state and len(st.session_state.selected_material_rolling) > 0:
            filtered_rolling = filtered_rolling[filtered_rolling[material_col_rolling].isin(st.session_state.selected_material_rolling)]
        
        # Try to find and filter by year week
        year_week_col = None
        for col in df_90_rolling.columns:
            if 'week' in str(col).lower() or 'fiscal' in str(col).lower():
                year_week_col = col
                break
        
        if year_week_col and 'selected_year_week' in st.session_state and len(st.session_state.selected_year_week) > 0:
            filtered_rolling = filtered_rolling[filtered_rolling[year_week_col].isin(st.session_state.selected_year_week)]
        
        # Display KPI cards
        st.subheader("Key Performance Indicators")
        
        kpi_cols = st.columns(5)
        kpi_metrics = {}
        
        # Try to extract KPI values
        for col in df_90_rolling.columns:
            col_lower = str(col).lower()
            if 'gdpw' in col_lower:
                kpi_metrics['GDPW'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'in house' in col_lower and 'production' in col_lower:
                kpi_metrics['In House Prod'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
            elif 'order' in col_lower and 'qty' not in col_lower:
                kpi_metrics['Orders'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'order' in col_lower and 'qty' in col_lower:
                kpi_metrics['Order Qty'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'received' in col_lower and 'good' in col_lower:
                kpi_metrics['Goods Received'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'yield' in col_lower and 'confirm' in col_lower and 'actual' in col_lower:
                kpi_metrics['Yield Confirmed'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'scrap' in col_lower and 'confirm' in col_lower:
                kpi_metrics['Scrap Confirmed'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'actual' in col_lower and 'yield' in col_lower and '%' in col:
                kpi_metrics['Avg Actual Yield %'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
            elif 'routing' in col_lower and 'yield' in col_lower:
                kpi_metrics['Avg Routing Yield %'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
            elif 'mrp' in col_lower and 'yield' in col_lower:
                kpi_metrics['Avg MRP Yield %'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
        
        # Display KPIs
        kpi_list = list(kpi_metrics.items())
        
        for idx, (kpi_name, kpi_value) in enumerate(kpi_list[:5]):
            with kpi_cols[idx % 5]:
                if isinstance(kpi_value, (int, float)):
                    if '%' in kpi_name:
                        st.metric(kpi_name, f"{kpi_value:.2f}%")
                    else:
                        st.metric(kpi_name, f"{int(kpi_value):,}")
        
        if len(kpi_list) > 5:
            kpi_cols2 = st.columns(min(5, len(kpi_list) - 5))
            for idx, (kpi_name, kpi_value) in enumerate(kpi_list[5:]):
                with kpi_cols2[idx]:
                    if isinstance(kpi_value, (int, float)):
                        if '%' in kpi_name:
                            st.metric(kpi_name, f"{kpi_value:.2f}%")
                        else:
                            st.metric(kpi_name, f"{int(kpi_value):,}")
        
        # Display table with conditional formatting
        st.subheader("90 Days Rolling History Data")
        st.dataframe(filtered_rolling, use_container_width=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in 90 Days Rolling History: {e}")

# ==================== SECTION 5: 90 DAY PIVOT ====================

st.header("5️⃣ 90 Day Pivot Summary")

if not df_90_pivot.empty:
    try:
        # Display pivot table
        st.subheader("Pivot Summary Table")
        st.dataframe(df_90_pivot, use_container_width=True)
        
        # Create comparison chart
        st.subheader("MRP Yield % vs Actual Yield % Comparison")
        
        # Try to find MRP Yield and Actual Yield columns
        mrp_yield_col = None
        actual_yield_col = None
        
        for col in df_90_pivot.columns:
            col_lower = str(col).lower()
            if 'mrp' in col_lower and 'yield' in col_lower and not mrp_yield_col:
                mrp_yield_col = col
            if 'actual' in col_lower and 'yield' in col_lower and not actual_yield_col:
                actual_yield_col = col
        
        if mrp_yield_col and actual_yield_col:
            pivot_chart_data = df_90_pivot.copy()
            pivot_chart_data[mrp_yield_col] = pd.to_numeric(pivot_chart_data[mrp_yield_col], errors='coerce')
            pivot_chart_data[actual_yield_col] = pd.to_numeric(pivot_chart_data[actual_yield_col], errors='coerce')
            
            # Get first column for x-axis (usually Material or Category)
            x_col = df_90_pivot.columns[0]
            
            fig_pivot = go.Figure(data=[
                go.Bar(name='MRP Yield %', x=pivot_chart_data[x_col], y=pivot_chart_data[mrp_yield_col]),
                go.Bar(name='Actual Yield %', x=pivot_chart_data[x_col], y=pivot_chart_data[actual_yield_col])
            ])
            
            fig_pivot.update_layout(
                title="Average MRP Yield % vs Average Actual Yield %",
                xaxis_title=x_col,
                yaxis_title="Yield %",
                barmode='group',
                hovermode='x unified',
                height=400
            )
            
            st.plotly_chart(fig_pivot, use_container_width=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in 90 Day Pivot: {e}")

# ==================== FOOTER ====================

st.markdown("""
---
### 📋 Dashboard Information
- **Last Updated**: Generated dynamically from Excel data
- **Data Source**: dummy_yield_history_data_updated.xlsx
- **All percentages are displayed as %**
""")
