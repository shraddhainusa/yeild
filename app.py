"""
Yield Performance Dashboard - Professional Version with Advanced Analytics
A comprehensive Streamlit dashboard with professional visuals, advanced analysis, and actionable insights
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
        
        /* Insight Card */
        .insight-card {
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            border-left: 4px solid #667eea;
            margin: 10px 0;
        }
        
        .insight-label {
            font-size: 11px;
            font-weight: 700;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }
        
        .insight-value {
            font-size: 28px;
            font-weight: 900;
            color: #1e3a8a;
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
        
        /* Management Insights */
        .management-insight {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            margin: 20px 0;
        }
        
        .insight-bullet {
            font-size: 15px;
            line-height: 1.8;
            margin: 10px 0;
            padding-left: 20px;
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

# ==================== HELPER FUNCTIONS ====================
def format_percentage(value):
    """Convert decimal to percentage format (0.93 -> 93%)"""
    if pd.isna(value):
        return 0
    if isinstance(value, str):
        try:
            value = float(value)
        except:
            return 0
    
    # If value is between 0 and 1, multiply by 100
    if 0 <= value <= 1:
        return value * 100
    # Otherwise assume it's already in percentage
    return value

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
    
    # Format percentages
    for col in ['MRP Plan', 'Actual', 'Variance']:
        if col in df_reshaped.columns:
            df_reshaped[col] = df_reshaped[col].apply(format_percentage)
    
    return df_reshaped

def enrich_overall_data(df):
    """Add calculated columns to overall data"""
    df_enriched = df.copy()
    
    # Find column names
    material_col = df_enriched.columns[0]
    mrp_col = None
    actual_col = None
    
    for col in df_enriched.columns:
        col_lower = str(col).lower()
        if 'mrp' in col_lower and 'plan' in col_lower and not mrp_col:
            mrp_col = col
        if 'actual' in col_lower and '%' in str(col) and not actual_col:
            actual_col = col
    
    if mrp_col and actual_col:
        # Format percentages
        df_enriched[mrp_col] = df_enriched[mrp_col].apply(format_percentage)
        df_enriched[actual_col] = df_enriched[actual_col].apply(format_percentage)
        
        # Calculate Variance
        df_enriched['Variance'] = df_enriched[actual_col] - df_enriched[mrp_col]
        
        # Risk Category
        def assign_risk(row):
            actual = format_percentage(row[actual_col])
            mrp = format_percentage(row[mrp_col])
            
            # Find order quantity
            order_qty = 0
            for col in df_enriched.columns:
                if 'order' in str(col).lower() and 'qty' in str(col).lower():
                    try:
                        order_qty = float(row[col])
                    except:
                        pass
                    break
            
            if actual < mrp:
                if order_qty > df_enriched[[col for col in df_enriched.columns 
                                           if 'order' in str(col).lower() and 'qty' in str(col).lower()][0]].median() if any('order' in str(col).lower() and 'qty' in str(col).lower() for col in df_enriched.columns) else 0:
                    return "High Risk"
                else:
                    return "Medium Risk"
            elif actual > mrp:
                return "Opportunity"
            else:
                return "Stable"
        
        df_enriched['Risk Category'] = df_enriched.apply(assign_risk, axis=1)
    
    return df_enriched

df_overall_enriched = enrich_overall_data(df_overall)
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
    if not df_overall_enriched.empty:
        materials_overall = df_overall_enriched[df_overall_enriched.columns[0]].unique()
        selected_material_overall = st.multiselect(
            "Select Materials",
            options=materials_overall,
            default=list(materials_overall)[:3] if len(materials_overall) > 0 else [],
            key="materials_overall"
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

st.sidebar.markdown("---")
st.sidebar.markdown("📌 **Tips:**")
st.sidebar.markdown("• Use filters to focus on specific data")
st.sidebar.markdown("• Hover over charts for details")
st.sidebar.markdown("• Click legend items to hide/show series")

# ==================== SECTION 1: EXECUTIVE SUMMARY ====================
st.markdown("<h2>1️⃣ Executive Summary</h2>", unsafe_allow_html=True)

if not df_overall_enriched.empty:
    try:
        col_names = df_overall_enriched.columns.tolist()
        metrics = {}
        
        # Calculate metrics with proper percentage formatting
        for col in col_names:
            col_lower = str(col).lower()
            if 'mrp' in col_lower and 'plan' in col_lower and '%' in str(col) and 'avg_mrp_plan' not in metrics:
                metrics['avg_mrp_plan'] = df_overall_enriched[col].apply(format_percentage).mean()
            if 'actual' in col_lower and '%' in str(col) and 'avg_actual' not in metrics:
                metrics['avg_actual'] = df_overall_enriched[col].apply(format_percentage).mean()
            if 'variance' in col_lower and 'avg_variance' not in metrics:
                metrics['avg_variance'] = df_overall_enriched['Variance'].mean()
            if 'order' in col_lower and 'qty' not in col_lower and 'total_orders' not in metrics:
                metrics['total_orders'] = pd.to_numeric(df_overall_enriched[col], errors='coerce').sum()
            if 'start' in col_lower and 'qty' in col_lower and 'start_qty' not in metrics:
                metrics['start_qty'] = pd.to_numeric(df_overall_enriched[col], errors='coerce').sum()
            if 'deliver' in col_lower and 'qty' in col_lower and 'delivered_qty' not in metrics:
                metrics['delivered_qty'] = pd.to_numeric(df_overall_enriched[col], errors='coerce').sum()
            if 'weighted' in col_lower and 'avg_weighted' not in metrics:
                metrics['avg_weighted'] = df_overall_enriched[col].apply(format_percentage).mean()
        
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
        
        # Second row of KPIs - FIXED Start & Delivered Qty
        kpi_col5, kpi_col6, kpi_col7 = st.columns(3)
        
        with kpi_col5:
            total_orders = metrics.get('total_orders', 0)
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #06b6d4 0%, #0891b2 100%);">
                <div class="kpi-label">📦 Total Orders</div>
                <div class="kpi-value">{int(total_orders):,}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with kpi_col6:
            start_qty = metrics.get('start_qty', 0)
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #ec4899 0%, #db2777 100%);">
                <div class="kpi-label">📤 Start Qty</div>
                <div class="kpi-value">{int(start_qty):,}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with kpi_col7:
            delivered_qty = metrics.get('delivered_qty', 0)
            st.markdown(f"""
            <div class="kpi-container" style="background: linear-gradient(135deg, #14b8a6 0%, #0d9488 100%);">
                <div class="kpi-label">📬 Delivered Qty</div>
                <div class="kpi-value">{int(delivered_qty):,}</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error calculating metrics: {e}")

# ==================== SECTION 1B: INSIGHTS ====================
st.markdown("<h2>📊 Key Insights</h2>", unsafe_allow_html=True)

if not df_overall_enriched.empty:
    try:
        # Find relevant columns
        material_col = df_overall_enriched.columns[0]
        actual_col = None
        mrp_col = None
        
        for col in df_overall_enriched.columns:
            col_lower = str(col).lower()
            if 'actual' in col_lower and '%' in str(col) and not actual_col:
                actual_col = col
            if 'mrp' in col_lower and 'plan' in col_lower and not mrp_col:
                mrp_col = col
        
        if actual_col and mrp_col:
            # Calculate insights
            df_overall_enriched[actual_col] = df_overall_enriched[actual_col].apply(format_percentage)
            df_overall_enriched[mrp_col] = df_overall_enriched[mrp_col].apply(format_percentage)
            
            below_mrp = len(df_overall_enriched[df_overall_enriched[actual_col] < df_overall_enriched[mrp_col]])
            above_mrp = len(df_overall_enriched[df_overall_enriched[actual_col] >= df_overall_enriched[mrp_col]])
            
            # Top negative variance
            df_overall_enriched['Variance'] = df_overall_enriched[actual_col] - df_overall_enriched[mrp_col]
            top_negative = df_overall_enriched.loc[df_overall_enriched['Variance'].idxmin()]
            top_positive = df_overall_enriched.loc[df_overall_enriched['Variance'].idxmax()]
            
            # Top scrap (if available)
            scrap_col = None
            for col in df_overall_enriched.columns:
                if 'scrap' in str(col).lower():
                    scrap_col = col
                    break
            
            top_scrap_material = "N/A"
            lowest_yield_material = "N/A"
            
            if scrap_col:
                top_scrap_idx = df_overall_enriched[scrap_col].idxmax()
                top_scrap_material = df_overall_enriched.loc[top_scrap_idx, material_col]
            
            lowest_yield_idx = df_overall_enriched[actual_col].idxmin()
            lowest_yield_material = df_overall_enriched.loc[lowest_yield_idx, material_col]
            
            # Display insight cards
            col_ins1, col_ins2, col_ins3, col_ins4 = st.columns(4)
            
            with col_ins1:
                st.markdown(f"""
                <div class="insight-card">
                    <div class="insight-label">📉 Below MRP</div>
                    <div class="insight-value">{below_mrp}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_ins2:
                st.markdown(f"""
                <div class="insight-card" style="border-left-color: #10b981;">
                    <div class="insight-label">📈 Above MRP</div>
                    <div class="insight-value">{above_mrp}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_ins3:
                st.markdown(f"""
                <div class="insight-card" style="border-left-color: #f59e0b;">
                    <div class="insight-label">⬇️ Top Negative</div>
                    <div class="insight-value">{top_negative[material_col]}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_ins4:
                st.markdown(f"""
                <div class="insight-card" style="border-left-color: #10b981;">
                    <div class="insight-label">⬆️ Top Positive</div>
                    <div class="insight-value">{top_positive[material_col]}</div>
                </div>
                """, unsafe_allow_html=True)
            
            col_ins5, col_ins6 = st.columns(2)
            
            with col_ins5:
                st.markdown(f"""
                <div class="insight-card" style="border-left-color: #ef4444;">
                    <div class="insight-label">🚨 Highest Scrap</div>
                    <div class="insight-value">{str(top_scrap_material)[:15]}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_ins6:
                st.markdown(f"""
                <div class="insight-card" style="border-left-color: #f59e0b;">
                    <div class="insight-label">📉 Lowest Yield</div>
                    <div class="insight-value">{str(lowest_yield_material)[:15]}</div>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.warning(f"Could not generate all insights: {e}")

# ==================== SECTION 2: YIELD HISTORY OVERALL ====================
st.markdown("<h2>2️⃣ Yield History Overall - Detailed Analysis</h2>", unsafe_allow_html=True)

if not df_overall_enriched.empty:
    try:
        filtered_df = df_overall_enriched.copy()
        
        if 'selected_material_overall' in st.session_state and len(st.session_state.selected_material_overall) > 0:
            material_col = df_overall_enriched.columns[0]
            filtered_df = filtered_df[filtered_df[material_col].isin(st.session_state.selected_material_overall)]
        
        # Create comparison chart
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.subheader("📊 Top 10 Negative Variance Materials")
            
            top_negative_df = df_overall_enriched.nsmallest(10, 'Variance')
            
            fig = go.Figure(data=[
                go.Bar(
                    x=top_negative_df[df_overall_enriched.columns[0]],
                    y=top_negative_df['Variance'],
                    marker_color='#ef4444',
                    marker_line_color='#dc2626',
                    marker_line_width=1.5
                )
            ])
            
            fig.update_layout(
                title="",
                xaxis_title="Material",
                yaxis_title="Variance %",
                hovermode='x unified',
                height=400,
                template='plotly_white',
                font=dict(size=11),
                margin=dict(b=80),
                showlegend=False
            )
            
            fig.update_xaxes(tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        
        with col_chart2:
            st.subheader("📈 Top 10 Positive Variance Materials")
            
            top_positive_df = df_overall_enriched.nlargest(10, 'Variance')
            
            fig = go.Figure(data=[
                go.Bar(
                    x=top_positive_df[df_overall_enriched.columns[0]],
                    y=top_positive_df['Variance'],
                    marker_color='#10b981',
                    marker_line_color='#059669',
                    marker_line_width=1.5
                )
            ])
            
            fig.update_layout(
                title="",
                xaxis_title="Material",
                yaxis_title="Variance %",
                hovermode='x unified',
                height=400,
                template='plotly_white',
                font=dict(size=11),
                margin=dict(b=80),
                showlegend=False
            )
            
            fig.update_xaxes(tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        
        # MRP Plan vs Actual chart
        st.subheader("🎯 MRP Plan vs Actual Yield % - All Materials")
        
        mrp_col = None
        actual_col = None
        material_col = df_overall_enriched.columns[0]
        
        for col in df_overall_enriched.columns:
            if 'mrp' in str(col).lower() and 'plan' in str(col).lower() and not mrp_col:
                mrp_col = col
            if 'actual' in str(col).lower() and '%' in str(col) and not actual_col:
                actual_col = col
        
        if mrp_col and actual_col:
            # Sort by variance for better visualization
            chart_data = df_overall_enriched[[material_col, mrp_col, actual_col]].copy().sort_values('Variance')
            
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
        
        # Data table
        st.subheader("📋 Detailed Yield Data with Risk Assessment")
        
        display_df = df_overall_enriched.copy()
        
        # Format for display
        for col in display_df.columns:
            if '%' in str(col) or col in ['Variance']:
                try:
                    display_df[col] = display_df[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A")
                except:
                    pass
        
        st.dataframe(display_df, use_container_width=True, height=400)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in Yield History Overall: {e}")

# ==================== SECTION 3: MONTHLY TRENDS ====================
st.markdown("<h2>3️⃣ Monthly Yield Trend Analysis</h2>", unsafe_allow_html=True)

if not df_monthly_long.empty:
    try:
        filtered_monthly = df_monthly_long.copy()
        
        if 'selected_material_monthly' in st.session_state and len(st.session_state.selected_material_monthly) > 0:
            filtered_monthly = filtered_monthly[filtered_monthly['Material'].isin(st.session_state.selected_material_monthly)]
        
        col_trend1, col_trend2 = st.columns(2)
        
        # Line chart for MRP Plan vs Actual
        with col_trend1:
            st.subheader("📈 Monthly MRP Plan vs Actual Trend")
            
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
            st.subheader("📊 Monthly Variance Trend")
            
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

# ==================== SECTION 4: 90 DAY ADVANCED ANALYSIS ====================
st.markdown("<h2>4️⃣ 90 Days Rolling History - Advanced Analytics</h2>", unsafe_allow_html=True)

if not df_90_rolling.empty:
    try:
        filtered_rolling = df_90_rolling.copy()
        
        # Find relevant columns
        material_col_rolling = None
        for col in df_90_rolling.columns:
            if 'material' in str(col).lower() or col == df_90_rolling.columns[0]:
                material_col_rolling = col
                break
        
        # Calculate Order Quantity vs Actual Yield scatter
        col_chart5, col_chart6 = st.columns(2)
        
        with col_chart5:
            st.write("**Actual Yield % vs Total Order Quantity**")
            
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
                scatter_data[actual_yield_col] = pd.to_numeric(scatter_data[actual_yield_col], errors='coerce').apply(format_percentage)
                
                fig_scatter = go.Figure(data=go.Scatter(
                    x=scatter_data[order_qty_col],
                    y=scatter_data[actual_yield_col],
                    mode='markers',
                    marker=dict(
                        size=10,
                        color=scatter_data[actual_yield_col],
                        colorscale='RdYlGn',
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
        
        with col_chart6:
            st.write("**Scrap Confirmed vs Actual Yield %**")
            
            scrap_col = None
            actual_yield_col = None
            
            for col in df_90_rolling.columns:
                col_lower = str(col).lower()
                if 'scrap' in col_lower and 'confirm' in col_lower and not scrap_col:
                    scrap_col = col
                if 'actual' in col_lower and 'yield' in col_lower and '%' in str(col) and not actual_yield_col:
                    actual_yield_col = col
            
            if scrap_col and actual_yield_col:
                scatter_data2 = filtered_rolling[[scrap_col, actual_yield_col]].copy()
                scatter_data2[scrap_col] = pd.to_numeric(scatter_data2[scrap_col], errors='coerce')
                scatter_data2[actual_yield_col] = pd.to_numeric(scatter_data2[actual_yield_col], errors='coerce').apply(format_percentage)
                
                fig_scatter2 = go.Figure(data=go.Scatter(
                    x=scatter_data2[scrap_col],
                    y=scatter_data2[actual_yield_col],
                    mode='markers',
                    marker=dict(
                        size=10,
                        color=scatter_data2[actual_yield_col],
                        colorscale='RdYlGn',
                        showscale=True,
                        line_width=1.5
                    )
                ))
                
                fig_scatter2.update_layout(
                    title="",
                    xaxis_title="Scrap Confirmed",
                    yaxis_title="Actual Yield %",
                    height=350,
                    template='plotly_white'
                )
                
                st.plotly_chart(fig_scatter2, use_container_width=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in 90 Day Analysis: {e}")

# ==================== SECTION 5: TOP UNDERPERFORMING MATERIALS ====================
st.markdown("<h2>5️⃣ Top 15 Underperforming vs Overperforming Materials</h2>", unsafe_allow_html=True)

if not df_90_pivot.empty:
    try:
        df_90_pivot_enriched = df_90_pivot.copy()
        
        # Find yield columns
        mrp_yield_col = None
        actual_yield_col = None
        
        for col in df_90_pivot_enriched.columns:
            col_lower = str(col).lower()
            if 'mrp' in col_lower and 'yield' in col_lower and not mrp_yield_col:
                mrp_yield_col = col
            if 'actual' in col_lower and 'yield' in col_lower and not actual_yield_col:
                actual_yield_col = col
        
        if mrp_yield_col and actual_yield_col:
            # Format percentages
            df_90_pivot_enriched[mrp_yield_col] = df_90_pivot_enriched[mrp_yield_col].apply(format_percentage)
            df_90_pivot_enriched[actual_yield_col] = df_90_pivot_enriched[actual_yield_col].apply(format_percentage)
            
            # Calculate gap
            df_90_pivot_enriched['Yield Gap'] = df_90_pivot_enriched[actual_yield_col] - df_90_pivot_enriched[mrp_yield_col]
            
            x_col = df_90_pivot_enriched.columns[0]
            
            # Top 15 underperforming
            underperforming = df_90_pivot_enriched.nsmallest(15, 'Yield Gap')
            # Top 15 overperforming
            overperforming = df_90_pivot_enriched.nlargest(15, 'Yield Gap')
            
            col_perf1, col_perf2 = st.columns(2)
            
            with col_perf1:
                st.subheader("📉 Top 15 Underperforming Materials")
                
                fig_under = go.Figure(data=[
                    go.Bar(
                        x=underperforming[x_col],
                        y=underperforming['Yield Gap'],
                        marker_color='#ef4444',
                        marker_line_color='#dc2626',
                        marker_line_width=1.5
                    )
                ])
                
                fig_under.update_layout(
                    title="",
                    xaxis_title="Material",
                    yaxis_title="Yield Gap %",
                    height=400,
                    template='plotly_white',
                    font=dict(size=10),
                    margin=dict(b=100),
                    showlegend=False
                )
                
                fig_under.update_xaxes(tickangle=-45)
                st.plotly_chart(fig_under, use_container_width=True)
            
            with col_perf2:
                st.subheader("📈 Top 15 Overperforming Materials")
                
                fig_over = go.Figure(data=[
                    go.Bar(
                        x=overperforming[x_col],
                        y=overperforming['Yield Gap'],
                        marker_color='#10b981',
                        marker_line_color='#059669',
                        marker_line_width=1.5
                    )
                ])
                
                fig_over.update_layout(
                    title="",
                    xaxis_title="Material",
                    yaxis_title="Yield Gap %",
                    height=400,
                    template='plotly_white',
                    font=dict(size=10),
                    margin=dict(b=100),
                    showlegend=False
                )
                
                fig_over.update_xaxes(tickangle=-45)
                st.plotly_chart(fig_over, use_container_width=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error in 90 Day Pivot: {e}")

# ==================== SECTION 6: MANAGEMENT INSIGHTS ====================
st.markdown("<h2>6️⃣ Management Insights & Recommendations</h2>", unsafe_allow_html=True)

if not df_overall_enriched.empty:
    try:
        # Generate insights
        material_col = df_overall_enriched.columns[0]
        actual_col = None
        mrp_col = None
        
        for col in df_overall_enriched.columns:
            col_lower = str(col).lower()
            if 'actual' in col_lower and '%' in str(col) and not actual_col:
                actual_col = col
            if 'mrp' in col_lower and 'plan' in col_lower and not mrp_col:
                mrp_col = col
        
        insights = []
        
        if actual_col and mrp_col:
            avg_actual = df_overall_enriched[actual_col].apply(format_percentage).mean()
            avg_mrp = df_overall_enriched[mrp_col].apply(format_percentage).mean()
            
            # Insight 1: Overall performance
            if avg_actual > avg_mrp:
                insights.append(f"✅ Overall Performance: Actual yield ({avg_actual:.1f}%) is EXCEEDING MRP plan ({avg_mrp:.1f}%). This indicates strong operational performance.")
            else:
                diff = avg_mrp - avg_actual
                insights.append(f"⚠️ Overall Performance: Actual yield ({avg_actual:.1f}%) is BELOW MRP plan ({avg_mrp:.1f}%) by {diff:.1f} percentage points. Immediate attention required.")
            
            # Insight 2: Materials needing attention
            underperformers = df_overall_enriched[df_overall_enriched[actual_col] < df_overall_enriched[mrp_col]]
            if len(underperformers) > 0:
                worst_material = underperformers.loc[underperformers[actual_col].idxmin()]
                insights.append(f"🎯 Critical Focus: {worst_material[material_col]} requires immediate intervention with actual yield at {worst_material[actual_col]:.1f}% vs plan of {worst_material[mrp_col]:.1f}%.")
            
            # Insight 3: Volume impact
            total_orders = 0
            order_qty_col = None
            for col in df_overall_enriched.columns:
                if 'order' in str(col).lower() and 'qty' in str(col).lower():
                    order_qty_col = col
                    break
            
            if order_qty_col:
                total_orders = df_overall_enriched[order_qty_col].sum()
                avg_order = df_overall_enriched[order_qty_col].mean()
                
                high_volume_underperformers = underperformers[underperformers[order_qty_col] > avg_order]
                if len(high_volume_underperformers) > 0:
                    insights.append(f"📦 Volume Impact: {len(high_volume_underperformers)} high-volume materials are underperforming. This has significant business impact as these represent {high_volume_underperformers[order_qty_col].sum()/total_orders*100:.1f}% of total order quantity.")
            
            # Insight 4: Planning review
            overperformers = df_overall_enriched[df_overall_enriched[actual_col] > df_overall_enriched[mrp_col]]
            if len(overperformers) / len(df_overall_enriched) > 0.3:
                insights.append(f"📊 Planning Assumptions: {len(overperformers)} materials ({len(overperformers)/len(df_overall_enriched)*100:.0f}%) are exceeding their plans consistently, suggesting MRP plans may be conservative and could be optimized.")
            
            # Insight 5: Scrap analysis
            scrap_col = None
            for col in df_overall_enriched.columns:
                if 'scrap' in str(col).lower():
                    scrap_col = col
                    break
            
            if scrap_col:
                top_scrap = df_overall_enriched[scrap_col].max()
                if top_scrap > 0:
                    scrap_contributor = df_overall_enriched.loc[df_overall_enriched[scrap_col].idxmax(), material_col]
                    insights.append(f"🚨 Scrap Contribution: {scrap_contributor} has the highest scrap confirmed. Investigating scrap reasons should be a priority as this directly impacts yield performance.")
        
        # Display insights
        for i, insight in enumerate(insights, 1):
            st.markdown(f"""
            <div class="management-insight">
                <div class="insight-bullet">• {insight}</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
    
    except Exception as e:
        st.warning(f"Could not generate all management insights: {e}")

# ==================== FOOTER ====================
st.markdown("""
<div style="text-align: center; padding: 40px 0; color: #666; border-top: 2px solid #e5e7eb; margin-top: 40px;">
    <p style="margin: 0; font-size: 12px;">
        📊 <b>Yield Performance Dashboard</b> | Professional Analytics Version
    </p>
    <p style="margin: 10px 0; font-size: 11px; color: #999;">
        Data Source: dummy_yield_history_data_updated.xlsx | Last Updated: Generated Dynamically
    </p>
</div>
""", unsafe_allow_html=True)
