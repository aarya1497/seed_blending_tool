import streamlit as st
import pandas as pd
import numpy as np
import io
import warnings
from scipy.optimize import linprog

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.set_page_config(
    page_title="Seed Blending & Inventory Optimization System", 
    page_icon="🌱", 
    layout="wide"
)

st.markdown("""
    <style>
    .main-header {
        font-size: 28px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 4px;
    }
    .sub-header {
        font-size: 15px;
        color: #4B5563;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🌱 Seed Blending & Inventory Allocation System</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automated Minimum-Labor Inventory Optimization Engine with Weighted Quality Specification Matching</div>', unsafe_allow_html=True)

@st.cache_data
def generate_sample_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sample_upload = pd.DataFrame({
            'Batch': ['KR237310', 'KR237378', 'KR237379', 'KR237380', 'KR237309', 'KR237319', 'RF7543N63', 'WR2120594'],
            'Category': ['FRS', 'FRS', 'FRS', 'FRS', 'FRS', 'FRS', 'Inventory', 'Inventory'],
            'SAP Total': [626, 957, 1720, 1196, 582, 1425, 1050, 5970],
            'OT-2% ': [5.95, 0.68, 0.22, 0.47, 5.65, 0.22, 6.08, 1.00],
            'G.P %': [93.59, 98.42, 98.21, 98.36, 93.65, 99.35, 93.92, 99.00],
            'Phy.Pty': ['OK', 'OK', 'OK', 'OK', 'OK', 'OK', 'OK', 'OK'],
            'First Count': [84, 80, 64, 84, 82, 82, 90, 80],
            'Final G%': [82, 87, 86, 85, 81, 82, 86, 82],
            'DOT': ['2024-01-09', '2024-01-30', '2024-01-30', '2024-01-09', '2024-01-09', '2024-01-09', '2024-01-18', '2024-01-19'],
            'Grade': ['C', 'C', 'C', 'A', 'C', 'B', 'C', 'A'],
            'location': ['KOTHURU COLD', 'KOTHURU COLD', 'KOTHURU COLD', 'KOTHURU COLD', 'KOTHURU COLD', 'KOTHURU COLD', 'BVN COLD', 'BVN COLD'],
            'Chamber': ['CHAMBER-2', 'CHAMBER-2', 'CHAMBER-2', 'CHAMBER-2', 'CHAMBER-2', 'CHAMBER-2', 'CHAMBER1', 'CHAMBER1'],
            'Floor': ['F1', 'F2', 'F2', 'F2', 'F1', 'F1', 'F-1', 'F-1'],
            'Bin': ['D1', 'A6', 'A6', 'A6', 'D1', 'D1', 'B-20', 'D-3']
        })
        
        pd.DataFrame().to_excel(writer, sheet_name='Paddy_Data_Upload', index=False)
        sample_upload.to_excel(writer, sheet_name='Paddy_Data_Upload', startrow=2, index=False)
        
        dashboard_meta = pd.DataFrame([
            ['Batch ID', 'Location', 'Chamber-Bin', 'SAP Total', 'Allocation Qty', 'OT-2', 'GP%', 'FC', 'Germ'],
            ['KR237310', 'KOTHURU COLD', 'CHAMBER-2-D1', 626, 0, 0, 0, 0, 0],
            ['KR237378', 'KOTHURU COLD', 'CHAMBER-2-A6', 957, 0, 0, 0, 0, 0],
            ['KR237379', 'KOTHURU COLD', 'CHAMBER-2-A6', 1720, 0, 0, 0, 0, 0],
            ['KR237380', 'KOTHURU COLD', 'CHAMBER-2-A6', 1196, 0, 0, 0, 0, 0],
            ['KR237309', 'KOTHURU COLD', 'CHAMBER-2-D1', 582, 0, 0, 0, 0, 0],
            ['KR237319', 'KOTHURU COLD', 'CHAMBER-2-D1', 1425, 0, 0, 0, 0, 0]
        ])
        dashboard_meta.to_excel(writer, sheet_name='Paddy_Blending_Dashboard', startrow=6, index=False, header=False)
        
    return output.getvalue()

def get_labor_rank(location, bin_code):
    loc = str(location).upper()
    b_code = str(bin_code).upper()
    if "KOTHURU" in loc:
        if "D1" in b_code:
            return 1
        elif "A6" in b_code:
            return 2
        else:
            return 2.5
    elif "BVN" in loc:
        if "B" in b_code:
            return 3
        elif "C" in b_code:
            return 4
        elif "D" in b_code:
            return 5
        else:
            return 6
    return 7

st.sidebar.header("🕹️ Executive Controls")

try:
    template_bytes = generate_sample_template()
    st.sidebar.download_button(
        label="📄 Download Input Excel Template",
        data=template_bytes,
        file_name="Seed_Blending_Data_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_template_dl"
    )
except Exception as e:
    st.sidebar.error(f"Template error: {e}")

st.sidebar.divider()

uploaded_file = st.sidebar.file_uploader(
    "Upload New Inventory File (.xlsx)", 
    type=["xlsx"],
    key="file_uploader"
)

location_option = st.sidebar.selectbox(
    "Location Strategy",
    (
        "1. All Locations (Smart Optimization - Least Labor)",
        "2. Force Only KOTHURU Floor 1 (D1)",
        "3. Force Only KOTHURU Floor 2 (A6)",
        "4. Force Only BVN COLD"
    )
)

target_qty = st.sidebar.number_input(
    "Target Demand Needed (kg)",
    min_value=1.0,
    max_value=100000.0,
    value=8595.0,
    step=250.0
)

max_ot2 = st.sidebar.slider(
    "Max Allowable OT-2 % Limit",
    min_value=0.5,
    max_value=5.0,
    value=1.5,
    step=0.1
)

def run_optimization(file_obj, target_qty, loc_option, max_ot2):
    if file_obj is None:
        try:
            file_obj = "Seed Blending Tool 2 (2).xlsx"
            xls = pd.ExcelFile(file_obj, engine='openpyxl')
        except Exception:
            return None, "Default dataset 'Seed Blending Tool 2 (2).xlsx' not found in current directory. Please upload an Excel file using the sidebar.", None
    else:
        file_bytes = file_obj.read()
        file_obj.seek(0)
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')

    try:
        df_upload = pd.read_excel(xls, sheet_name='Paddy_Data_Upload', skiprows=2)
    except Exception:
        df_upload = pd.read_excel(xls, sheet_name=0, skiprows=2)

    col_map = {c: str(c).strip() for c in df_upload.columns}
    df_upload = df_upload.rename(columns=col_map)
    if 'OT-2%' not in df_upload.columns and 'OT-2% ' in df_upload.columns:
        df_upload = df_upload.rename(columns={'OT-2% ': 'OT-2%'})

    if 'Batch' not in df_upload.columns and 'Batch ID' in df_upload.columns:
        df_upload = df_upload.rename(columns={'Batch ID': 'Batch'})
    
    df_upload = df_upload.dropna(subset=['Batch'])
    
    if 'location' not in df_upload.columns:
        df_upload['location'] = 'KOTHURU COLD'
    if 'Chamber' not in df_upload.columns:
        df_upload['Chamber'] = 'CHAMBER-2'
    if 'Bin' not in df_upload.columns:
        df_upload['Bin'] = 'D1'

    df_upload['Location_upload'] = df_upload['location']
    df_upload['Chamber-Bin_upload'] = df_upload['Chamber'].astype(str) + '-' + df_upload['Bin'].astype(str)

    try:
        dashboard_raw = pd.read_excel(xls, sheet_name='Paddy_Blending_Dashboard', header=None)
        dash_rows = dashboard_raw.iloc[6:].copy()
        dash_rows.columns = [str(c).strip() for c in dash_rows.iloc[0]]
        dash_rows = dash_rows.iloc[1:].dropna(subset=['Batch ID'])
    except Exception:
        dash_rows = pd.DataFrame(columns=['Batch ID', 'Location', 'Chamber-Bin'])

    df_merged = pd.merge(
        df_upload,
        dash_rows[['Batch ID', 'Location', 'Chamber-Bin']], 
        left_on='Batch', right_on='Batch ID', how='left',
        suffixes=('_upload', '_dash')
    )
    
    if 'Location_dash' in df_merged.columns:
        df_merged['Location'] = df_merged['Location_dash'].fillna(df_merged['Location_upload'])
    else:
        df_merged['Location'] = df_merged['Location_upload']

    if 'Chamber-Bin_dash' in df_merged.columns:
        df_merged['Chamber-Bin'] = df_merged['Chamber-Bin_dash'].fillna(df_merged['Chamber-Bin_upload'])
    else:
        df_merged['Chamber-Bin'] = df_merged['Chamber-Bin_upload']

    df_merged['Batch ID'] = df_merged['Batch']
    df_merged['Labor_Rank'] = df_merged.apply(lambda r: get_labor_rank(r['Location'], r['Chamber-Bin']), axis=1)
    
    if "2." in loc_option:
        df_filtered = df_merged[df_merged['Chamber-Bin'].str.contains('D1', case=False, na=False)].copy()
    elif "3." in loc_option:
        df_filtered = df_merged[df_merged['Chamber-Bin'].str.contains('A6', case=False, na=False)].copy()
    elif "4." in loc_option:
        df_filtered = df_merged[df_merged['Location'].str.contains('BVN', case=False, na=False)].copy()
    else:
        df_filtered = df_merged.copy()

    if df_filtered.empty:
        return None, "No inventory matches the selected location filter.", None

    n = len(df_filtered)
    c = df_filtered['Labor_Rank'].values
    A_ub = [(df_filtered['OT-2%'].values - max_ot2)]
    b_ub = [0.0]
    A_eq = [np.ones(n)]
    b_eq = [target_qty]
    bounds = [(0, stock) for stock in df_filtered['SAP Total'].values]

    res = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method='highs')
    
    df_res = df_filtered.copy()
    if res.success:
        df_res['Allocation Qty'] = np.round(res.x, 2)
        status_text = "✅ Priority 1 (Optimized Blend Met Quality Limit)"
    else:
        df_res = df_res.sort_values(by=['Labor_Rank', 'OT-2%']).reset_index(drop=True)
        rem = target_qty
        allocated = []
        for _, r in df_res.iterrows():
            take = min(r['SAP Total'], max(0.0, rem))
            allocated.append(take)
            rem -= take
        df_res['Allocation Qty'] = allocated
        status_text = "⚠️ Limit Exceeded (Insufficient Clean Stock)"

    df_allocated = df_res[df_res['Allocation Qty'] > 0].copy()
    total_allocated = df_allocated['Allocation Qty'].sum()

    if total_allocated > 0:
        df_allocated['OT-2'] = np.round(df_allocated['Allocation Qty'] * df_allocated['OT-2%'], 2)
        df_allocated['GP%'] = np.round(df_allocated['Allocation Qty'] * df_allocated['G.P %'], 2)
        df_allocated['FC'] = np.round(df_allocated['Allocation Qty'] * df_allocated['First Count'], 2)
        df_allocated['Germ'] = np.round(df_allocated['Allocation Qty'] * df_allocated['Final G%'], 2)

        blend_ot2 = df_allocated['OT-2'].sum() / total_allocated
        blend_gp  = df_allocated['GP%'].sum() / total_allocated
        blend_fc  = df_allocated['FC'].sum() / total_allocated
        blend_germ = df_allocated['Germ'].sum() / total_allocated
    else:
        blend_ot2 = blend_gp = blend_fc = blend_germ = 0.0
        df_allocated['OT-2'] = df_allocated['GP%'] = df_allocated['FC'] = df_allocated['Germ'] = 0.0

    metrics = {
        'Target Needed': target_qty,
        'Total Allocated': total_allocated,
        'Blended OT-2 %': blend_ot2,
        'Blended GP %': blend_gp,
        'Blended FC %': blend_fc,
        'Blended Germ %': blend_germ,
        'Status': status_text
    }

    out_cols = ['Batch ID', 'Location', 'Chamber-Bin', 'Labor_Rank', 'SAP Total', 'Allocation Qty', 'OT-2', 'GP%', 'FC', 'Germ']
    available_cols = [col for col in out_cols if col in df_allocated.columns]
    return metrics, df_allocated[available_cols], None

metrics, df_allocated, err = run_optimization(uploaded_file, target_qty, location_option, max_ot2)

if err:
    st.warning(f"📌 Notice: {err}")
elif metrics:
    st.subheader("📊 Weighted Average Quality Indicators")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Target Needed", f"{metrics['Target Needed']:,.0f} kg")
    m2.metric("Total Allocated", f"{metrics['Total Allocated']:,.0f} kg")
    m3.metric("Weighted OT-2 %", f"{metrics['Blended OT-2 %']:.3f}%")
    m4.metric("Weighted GP %", f"{metrics['Blended GP %']:.2f}%")
    m5.metric("Weighted Germ %", f"{metrics['Blended Germ %']:.2f}%")

    if "✅" in metrics['Status']:
        st.success(f"Status: {metrics['Status']}")
    else:
        st.warning(f"Status: {metrics['Status']}")

    st.subheader("📋 Allocated Batches Output")
    st.dataframe(df_allocated, use_container_width=True)

    def create_result_excel(df_out, metrics_dict):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            summary_df = pd.DataFrame([
                ["Target Demand Needed (kg)", metrics_dict['Target Needed']],
                ["Total Stock Allocated (kg)", metrics_dict['Total Allocated']],
                ["Weighted Blended OT-2 %", metrics_dict['Blended OT-2 %']],
                ["Weighted Blended GP %", metrics_dict['Blended GP %']],
                ["Weighted Blended First Count %", metrics_dict['Blended FC %']],
                ["Weighted Blended Final Germination %", metrics_dict['Blended Germ %']],
                ["Optimization Status", metrics_dict['Status']]
            ], columns=["Metric", "Value"])
            summary_df.to_excel(writer, sheet_name="Executive_Summary", index=False)
            df_out.to_excel(writer, sheet_name="Allocated_Batches", index=False)
        return buffer.getvalue()

    st.subheader("📤 Export Reports")
    c1, c2 = st.columns(2)
    with c1:
        try:
            excel_report_bytes = create_result_excel(df_allocated, metrics)
            st.download_button(
                label="📊 Download Full Optimization Excel Report",
                data=excel_report_bytes,
                file_name="Optimized_Seed_Blending_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Generates an executive Excel workbook with KPI summaries and batch details.",
                key="btn_excel_dl"
            )
        except Exception as ex:
            st.error(f"Excel export error: {ex}")

    with c2:
        csv_data = df_allocated.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Download Allocation CSV",
            data=csv_data,
            file_name="Seed_Allocation_Batches.csv",
            mime="text/csv",
            key="btn_csv_dl"
        )