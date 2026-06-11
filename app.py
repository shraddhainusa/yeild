"""
Yield Performance Dashboard - Professional Version
A comprehensive Streamlit dashboard with beautiful visuals and interactive charts
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')

# ==================== PAGE CONFIGURATION ====================
st.set_page_config(
    page_title="Yield Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CUSTOM CSS STYLING ====================
st.markdown("""
    <style>
        /* Main background */
        .main {
            background-color: #f5f7fa;
        }
        
        /* Sidebar styling */
        .sidebar .sidebar-content {
            background-color: #1e3a8a;
            color: white;
        }
        
        /* KPI Card Styling */
        .kpi-container {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 8px 16px rgba(0,0,0,0.1);
            color: white;
            margin: 10px 0;
            border-left: 5px solid #fbbf24;
        }
        
        .kpi-label {
            font-size: 13px;
            font-weight: 600;
            opacity: 0.9;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .kpi-value {
            font-size: 42px;
            font-weight: 900;
            margin-top: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        
        /* Section Headers */
        .section-header {
            border-bottom: 3px solid #667eea;
            padding-bottom: 15px;
            margin: 30px 0 20px 0;
            color: #1e3a8a;
        }
        
        /* Cards with shadow */
        .metric-card {
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            border-left: 4px solid #667eea;
        }
        
        /* Chart container */
        .chart-container {
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            margin: 20px 0;
        }
        
        /* Table styling */
        .dataframe {
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        
        /* Title styling */
        h1 {
            color: #1e3a8a;
            text-align: center;
            margin-bottom: 30px;
            font-size: 3em;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }
        
        h2 {
            color: #1e3a8a;
            border-left: 5px solid #667eea;
            padding-left: 15px;
            margin-top: 30px;
        }
        
        h3 {
            color: #2d3748;
            margin-top: 20px;
        }
        
        /* Filter container */
        .filter-container {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            border-radius: 10px;
            color: white;
            margin: 20px 0;
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
    st.info("Please place the Excel file in the same directory as this script.")
    st.stop()

try:
    sheets = load_excel_data(data_file)
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()

# Get individual sheets
df_overall = sheets.get("Yield History Overall", pd.DataFrame())
df_monthly = sheets.get("Yield History Monthly", pd.DataFrame())
df_90_rolling = sheets.get("90 Days Rolling History", pd.DataFrame())
df_90_pivot = sheets.get("90 Day Pivot", pd.DataFrame())

# ==================== DATA RESHAPING ====================
def reshape_monthly_data(df):
    """Convert monthly sheet from wide to long format"""
    if df.empty:
        return df
    
    df_copy = df.copy()
    material_col = df_copy.columns[0]
    df_copy = df_copy.rename(columns={material_col: 'Material'})
    
    id_vars = ['Material']
    value_vars = [col for col in df_copy.columns if col != 'Material']
    
    df_melted = df_copy.melt(id_vars=id_vars, value_vars=value_vars, 
                              var_name='Period', value_name='Value')
    
    df_melted['Metric'] = df_melted['Period'].str.extract(r'(MRP Plan|Actual|Variance)')
    df_melted['Period_Date'] = df_melted['Period'].str.replace(r'(MRP Plan|Actual|Variance)', '', regex=True).str.strip()
    
    df_reshaped = df_melted.pivot_table(
        index=['Material', 'Period_Date'],
        columns='Metric',
        values='Value',
        aggfunc='first'
    ).reset_index()
    
    return df_reshaped

df_monthly_long = reshape_monthly_data(df_monthly)

# ==================== HEADER ====================
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("<h1>📊 Yield Performance Dashboard</h1>", unsafe_allow_html=True)

st.markdown("---")

# ==================== SIDEBAR FILTERS ====================
st.sidebar.markdown("## 🔧 Dashboard Controls")
st.sidebar.markdown("---")

# Section 1: Executive Summary
with st.sidebar.expander("📈 Executive Summary", expanded=True):
    st.write("**No filters needed**")
    st.info("View all metrics at a glance")

# Section 2: Yield History Overall
with st.sidebar.expander("📊 Yield History Overall", expanded=False):
    if not df_overall.empty:
        materials_overall = df_overall[df_overall.columns[0]].unique()
        selected_material_overall = st.multiselect(
            "Select Materials",
            options=materials_overall,
            default=list(materials_overall)[:3] if len(materials_overall) > 0 else [],
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
            default=list(materials_monthly)[:2] if len(materials_monthly) > 0 else [],
            key="materials_monthly"
        )

# Section 4: 90 Days Rolling History
with st.sidebar.expander("🎯 90 Days Rolling History", expanded=False):
    if not df_90_rolling.empty:
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
                default=list(materials_rolling)[:2] if len(materials_rolling) > 0 else [],
                key="materials_rolling"
            )
        
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

st.sidebar.markdown("---")
st.sidebar.markdown("📌 **Tips:**")
st.sidebar.markdown("• Use filters to focus on specific data")
st.sidebar.markdown("• Hover over charts for details")
st.sidebar.markdown("• Click legend items to hide/show series")

# ==================== SECTION 1: EXECUTIVE SUMMARY ====================
st.markdown("<h2>1️⃣ Executive Summary</h2>", unsafe_allow_html=True)

if not df_overall.empty:
    try:
        col_names = df_overall.columns.tolist()
        metrics = {}
        
        # Calculate metrics
        for col in col_names:
            col_lower = str(col).lower()
            if 'mrp' in col_lower and 'plan' in col_lower and 'avg_mrp_plan' not in metrics:
                metrics['avg_mrp_plan'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
            if 'actual' in col_lower and '%' in str(col) and 'avg_actual' not in metrics:
                metrics['avg_actual'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
            if 'variance' in col_lower and 'avg_variance' not in metrics:
                metrics['avg_variance'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
            if 'order' in col_lower and 'qty' not in col_lower and 'total_orders' not in metrics:
                metrics['total_orders'] = pd.to_numeric(df_overall[col], errors='coerce').sum()
            if 'start' in col_lower and 'qty' in col_lower and 'start_qty' not in metrics:
                metrics['start_qty'] = pd.to_numeric(df_overall[col], errors='coerce').sum()
            if 'deliver' in col_lower and 'qty' in col_lower and 'delivered_qty' not in metrics:
                metrics['delivered_qty'] = pd.to_numeric(df_overall[col], errors='coerce').sum()
            if 'weighted' in col_lower and 'avg_weighted' not in metrics:
                metrics['avg_weighted'] = pd.to_numeric(df_overall[col], errors='coerce').mean()
        
        # Display KPI cards in rows
        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
        
        with kpi_col1:
            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-label">📊 Avg MRP Plan %</div>
                <div class="kpi-value">{metrics.get('avg_mrp_plan', 0):.1f}%</div>
            </div>
            """, unsafe_allow_html=True)
        
        with kpi_col2:
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #10b981 0%, #059669 100%);">
                <div class="kpi-label">✅ Avg Actual %</div>
                <div class="kpi-value">{metrics.get('avg_actual', 0):.1f}%</div>
            </div>
            """, unsafe_allow_html=True)
        
        with kpi_col3:
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);">
                <div class="kpi-label">📈 Avg Variance %</div>
                <div class="kpi-value">{metrics.get('avg_variance', 0):.1f}%</div>
            </div>
            """, unsafe_allow_html=True)
        
        with kpi_col4:
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%);">
                <div class="kpi-label">⚖️ Avg Weighted %</div>
                <div class="kpi-value">{metrics.get('avg_weighted', 0):.1f}%</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Second row of KPIs
        kpi_col5, kpi_col6, kpi_col7 = st.columns(3)
        
        with kpi_col5:
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #06b6d4 0%, #0891b2 100%);">
                <div class="kpi-label">📦 Total Orders</div>
                <div class="kpi-value">{int(metrics.get('total_orders', 0)):,}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with kpi_col6:
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #ec4899 0%, #db2777 100%);">
                <div class="kpi-label">📤 Start Qty</div>
                <div class="kpi-value">{int(metrics.get('start_qty', 0)):,}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with kpi_col7:
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #14b8a6 0%, #0d9488 100%);">
                <div class="kpi-label">📬 Delivered Qty</div>
                <div class="kpi-value">{int(metrics.get('delivered_qty', 0)):,}</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error calculating metrics: {e}")

# ==================== SECTION 2: YIELD HISTORY OVERALL ====================
st.markdown("<h2>2️⃣ Yield History Overall</h2>", unsafe_allow_html=True)

if not df_overall.empty:
    try:
        filtered_df = df_overall.copy()
        
        if 'selected_material_overall' in st.session_state and len(st.session_state.selected_material_overall) > 0:
            material_col = df_overall.columns[0]
            filtered_df = filtered_df[filtered_df[material_col].isin(st.session_state.selected_material_overall)]
        
        # Create comparison chart
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.subheader("📊 MRP Plan vs Actual Yield %")
            
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
                    go.Bar(name='MRP Plan %', x=chart_data[material_col], y=chart_data[mrp_col],
                           marker_color='#667eea', marker_line_color='#5568d3', marker_line_width=1.5),
                    go.Bar(name='Actual %', x=chart_data[material_col], y=chart_data[actual_col],
                           marker_color='#10b981', marker_line_color='#059669', marker_line_width=1.5)
                ])
                
                fig.update_layout(
                    title="",
                    xaxis_title="Material",
                    yaxis_title="Yield %",
                    barmode='group',
                    hovermode='x unified',
                    height=400,
                    template='plotly_white',
                    font=dict(size=11),
                    margin=dict(b=80),
                    showlegend=True,
                    legend=dict(x=0.5, y=1.1, orientation="h")
                )
                
                fig.update_xaxes(tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        
        with col_chart2:
            st.subheader("📈 Variance Distribution")
            
            variance_col = None
            for col in df_overall.columns:
                if 'variance' in str(col).lower():
                    variance_col = col
                    break
            
            if variance_col:
                variance_data = pd.to_numeric(df_overall[variance_col], errors='coerce')
                
                fig_var = go.Figure(data=[
                    go.Histogram(x=variance_data, nbinsx=15, 
                                marker_color='#f59e0b', marker_line_color='#d97706', marker_line_width=1)
                ])
                
                fig_var.update_layout(
                    title="",
                    xaxis_title="Variance %",
                    yaxis_title="Frequency",
                    height=400,
                    template='plotly_white',
                    font=dict(size=11),
                    showlegend=False
                )
                
                st.plotly_chart(fig_var, use_container_width=True)
        
        # Data table
        st.subheader("📋 Detailed Yield Data")
        
        # Format the table for display
        display_df = filtered_df.copy()
        
        # Apply conditional styling
        def highlight_underperformance(val):
            if isinstance(val, (int, float)) and '%' in str(val):
                if val < 80:
                    return 'background-color: #fee2e2'
            return ''
        
        st.dataframe(display_df, use_container_width=True, height=400)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in Yield History Overall: {e}")

# ==================== SECTION 3: MONTHLY YIELD TREND ====================
st.markdown("<h2>3️⃣ Monthly Yield Trend Analysis</h2>", unsafe_allow_html=True)

if not df_monthly_long.empty:
    try:
        filtered_monthly = df_monthly_long.copy()
        
        if 'selected_material_monthly' in st.session_state and len(st.session_state.selected_material_monthly) > 0:
            filtered_monthly = filtered_monthly[filtered_monthly['Material'].isin(st.session_state.selected_material_monthly)]
        
        col_trend1, col_trend2 = st.columns(2)
        
        # Line chart for MRP Plan vs Actual
        with col_trend1:
            st.subheader("📈 MRP Plan vs Actual Over Time")
            
            if 'MRP Plan' in filtered_monthly.columns and 'Actual' in filtered_monthly.columns:
                fig_trend = go.Figure()
                
                colors = px.colors.qualitative.Set2
                material_list = filtered_monthly['Material'].unique()
                
                for idx, material in enumerate(material_list):
                    mat_data = filtered_monthly[filtered_monthly['Material'] == material].sort_values('Period_Date')
                    color = colors[idx % len(colors)]
                    
                    fig_trend.add_trace(go.Scatter(
                        x=mat_data['Period_Date'],
                        y=mat_data['MRP Plan'],
                        mode='lines+markers',
                        name=f'{material} - MRP Plan',
                        line=dict(color=color, width=2),
                        marker=dict(size=6)
                    ))
                    
                    fig_trend.add_trace(go.Scatter(
                        x=mat_data['Period_Date'],
                        y=mat_data['Actual'],
                        mode='lines+markers',
                        name=f'{material} - Actual',
                        line=dict(color=color, width=2, dash='dash'),
                        marker=dict(size=6, symbol='diamond')
                    ))
                
                fig_trend.update_layout(
                    title="",
                    xaxis_title="Period",
                    yaxis_title="Yield %",
                    hovermode='x unified',
                    height=400,
                    template='plotly_white',
                    font=dict(size=11)
                )
                
                st.plotly_chart(fig_trend, use_container_width=True)
        
        # Variance chart
        with col_trend2:
            st.subheader("📊 Variance Trend")
            
            if 'Variance' in filtered_monthly.columns:
                fig_var_trend = go.Figure()
                
                for idx, material in enumerate(filtered_monthly['Material'].unique()):
                    mat_data = filtered_monthly[filtered_monthly['Material'] == material].sort_values('Period_Date')
                    color = colors[idx % len(colors)]
                    
                    fig_var_trend.add_trace(go.Scatter(
                        x=mat_data['Period_Date'],
                        y=mat_data['Variance'],
                        mode='lines+markers',
                        name=material,
                        line=dict(color=color, width=2.5),
                        marker=dict(size=8),
                        fill='tozeroy'
                    ))
                
                fig_var_trend.update_layout(
                    title="",
                    xaxis_title="Period",
                    yaxis_title="Variance %",
                    hovermode='x unified',
                    height=400,
                    template='plotly_white',
                    font=dict(size=11)
                )
                
                st.plotly_chart(fig_var_trend, use_container_width=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in Monthly Yield Trend: {e}")

# ==================== SECTION 4: 90 DAYS ROLLING HISTORY ====================
st.markdown("<h2>4️⃣ 90 Days Rolling History</h2>", unsafe_allow_html=True)

if not df_90_rolling.empty:
    try:
        filtered_rolling = df_90_rolling.copy()
        
        material_col_rolling = None
        for col in df_90_rolling.columns:
            if 'material' in str(col).lower() or col == df_90_rolling.columns[0]:
                material_col_rolling = col
                break
        
        if material_col_rolling and 'selected_material_rolling' in st.session_state and len(st.session_state.selected_material_rolling) > 0:
            filtered_rolling = filtered_rolling[filtered_rolling[material_col_rolling].isin(st.session_state.selected_material_rolling)]
        
        year_week_col = None
        for col in df_90_rolling.columns:
            if 'week' in str(col).lower() or 'fiscal' in str(col).lower():
                year_week_col = col
                break
        
        if year_week_col and 'selected_year_week' in st.session_state and len(st.session_state.selected_year_week) > 0:
            filtered_rolling = filtered_rolling[filtered_rolling[year_week_col].isin(st.session_state.selected_year_week)]
        
        # Display KPI cards
        st.subheader("🎯 Key Performance Indicators")
        
        kpi_metrics = {}
        
        for col in df_90_rolling.columns:
            col_lower = str(col).lower()
            if 'gdpw' in col_lower and 'gdpw' not in kpi_metrics:
                kpi_metrics['GDPW'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'in house' in col_lower and 'production' in col_lower:
                kpi_metrics['In House Production'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
            elif 'order' in col_lower and 'qty' not in col_lower and 'orders' not in kpi_metrics:
                kpi_metrics['Total Orders'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'order' in col_lower and 'qty' in col_lower and 'order_qty' not in kpi_metrics:
                kpi_metrics['Order Qty'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'received' in col_lower and 'good' in col_lower:
                kpi_metrics['Goods Received'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'yield' in col_lower and 'confirm' in col_lower and 'actual' not in col_lower:
                kpi_metrics['Yield Confirmed'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'scrap' in col_lower and 'confirm' in col_lower:
                kpi_metrics['Scrap Confirmed'] = pd.to_numeric(df_90_rolling[col], errors='coerce').sum()
            elif 'actual' in col_lower and 'yield' in col_lower and '%' in str(col):
                kpi_metrics['Actual Yield %'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
            elif 'routing' in col_lower and 'yield' in col_lower:
                kpi_metrics['Routing Yield %'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
            elif 'mrp' in col_lower and 'yield' in col_lower:
                kpi_metrics['MRP Yield %'] = pd.to_numeric(df_90_rolling[col], errors='coerce').mean()
        
        kpi_list = list(kpi_metrics.items())
        
        # Display KPIs in grid
        cols_per_row = 5
        for i in range(0, len(kpi_list), cols_per_row):
            kpi_cols = st.columns(min(cols_per_row, len(kpi_list) - i))
            for idx, (kpi_name, kpi_value) in enumerate(kpi_list[i:i+cols_per_row]):
                with kpi_cols[idx]:
                    if isinstance(kpi_value, (int, float)):
                        if '%' in kpi_name:
                            value_str = f"{kpi_value:.1f}%"
                            color = "#10b981"
                        else:
                            value_str = f"{int(kpi_value):,}"
                            color = "#667eea"
                        
                        st.markdown(f"""
                        <div class="kpi-container" style="background: linear-gradient(135deg, {color} 0%, {color}dd 100%); padding: 15px;">
                            <div class="kpi-label" style="font-size: 11px;">{kpi_name}</div>
                            <div class="kpi-value" style="font-size: 28px;">{value_str}</div>
                        </div>
                        """, unsafe_allow_html=True)
        
        # Charts section
        st.subheader("📊 Performance Charts")
        
        col_chart3, col_chart4 = st.columns(2)
        
        with col_chart3:
            st.write("**Yield % Comparison**")
            
            yield_cols = {}
            for col in df_90_rolling.columns:
                col_lower = str(col).lower()
                if 'actual' in col_lower and 'yield' in col_lower and '%' in str(col):
                    yield_cols['Actual'] = col
                elif 'routing' in col_lower and 'yield' in col_lower:
                    yield_cols['Routing'] = col
                elif 'mrp' in col_lower and 'yield' in col_lower:
                    yield_cols['MRP'] = col
            
            if len(yield_cols) > 0:
                fig_yield = go.Figure()
                
                for name, col in yield_cols.items():
                    values = pd.to_numeric(filtered_rolling[col], errors='coerce')
                    fig_yield.add_trace(go.Box(y=values, name=name))
                
                fig_yield.update_layout(
                    title="",
                    yaxis_title="Yield %",
                    height=350,
                    template='plotly_white',
                    showlegend=True
                )
                
                st.plotly_chart(fig_yield, use_container_width=True)
        
        with col_chart4:
            st.write("**Order Quantity vs Actual Yield %**")
            
            order_qty_col = None
            actual_yield_col = None
            
            for col in df_90_rolling.columns:
                col_lower = str(col).lower()
                if 'order' in col_lower and 'qty' in col_lower and not order_qty_col:
                    order_qty_col = col
                if 'actual' in col_lower and 'yield' in col_lower and '%' in str(col) and not actual_yield_col:
                    actual_yield_col = col
            
            if order_qty_col and actual_yield_col:
                scatter_data = filtered_rolling[[order_qty_col, actual_yield_col]].copy()
                scatter_data[order_qty_col] = pd.to_numeric(scatter_data[order_qty_col], errors='coerce')
                scatter_data[actual_yield_col] = pd.to_numeric(scatter_data[actual_yield_col], errors='coerce')
                
                fig_scatter = go.Figure(data=go.Scatter(
                    x=scatter_data[order_qty_col],
                    y=scatter_data[actual_yield_col],
                    mode='markers',
                    marker=dict(
                        size=8,
                        color=scatter_data[actual_yield_col],
                        colorscale='Viridis',
                        showscale=True,
                        line_width=1.5
                    )
                ))
                
                fig_scatter.update_layout(
                    title="",
                    xaxis_title="Order Quantity",
                    yaxis_title="Actual Yield %",
                    height=350,
                    template='plotly_white'
                )
                
                st.plotly_chart(fig_scatter, use_container_width=True)
        
        # Data table
        st.subheader("📋 Rolling History Data")
        st.dataframe(filtered_rolling, use_container_width=True, height=400)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in 90 Days Rolling History: {e}")

# ==================== SECTION 5: 90 DAY PIVOT ====================
st.markdown("<h2>5️⃣ 90 Day Pivot Summary</h2>", unsafe_allow_html=True)

if not df_90_pivot.empty:
    try:
        st.subheader("📊 Comparison Chart")
        
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
            
            x_col = df_90_pivot.columns[0]
            
            fig_pivot = go.Figure(data=[
                go.Bar(name='MRP Yield %', x=pivot_chart_data[x_col], y=pivot_chart_data[mrp_yield_col],
                       marker_color='#667eea', marker_line_color='#5568d3', marker_line_width=1.5),
                go.Bar(name='Actual Yield %', x=pivot_chart_data[x_col], y=pivot_chart_data[actual_yield_col],
                       marker_color='#10b981', marker_line_color='#059669', marker_line_width=1.5)
            ])
            
            fig_pivot.update_layout(
                title="",
                xaxis_title=x_col,
                yaxis_title="Yield %",
                barmode='group',
                hovermode='x unified',
                height=400,
                template='plotly_white',
                font=dict(size=11),
                margin=dict(b=80),
                showlegend=True,
                legend=dict(x=0.5, y=1.1, orientation="h")
            )
            
            fig_pivot.update_xaxes(tickangle=-45)
            st.plotly_chart(fig_pivot, use_container_width=True)
        
        st.subheader("📋 Pivot Summary Table")
        st.dataframe(df_90_pivot, use_container_width=True, height=400)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in 90 Day Pivot: {e}")

# ==================== FOOTER ====================
st.markdown("""
<div style="text-align: center; padding: 40px 0; color: #666; border-top: 2px solid #e5e7eb; margin-top: 40px;">
    <p style="margin: 0; font-size: 12px;">
        📊 <b>Yield Performance Dashboard</b> | Last Updated: Generated Dynamically from Excel Data
    </p>
    <p style="margin: 10px 0; font-size: 11px; color: #999;">
        Data Source: dummy_yield_history_data_updated.xlsx
    </p>
</div>
""", unsafe_allow_html=True)
