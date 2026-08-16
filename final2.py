import streamlit as st
import pandas as pd
import numpy as np
import io
import math
import warnings
from scipy.optimize import milp, LinearConstraint, Bounds

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
        color: #9CA3AF;
        margin-bottom: 20px;
    }
    .kpi-container {
        background-color: #1E293B;
        border-radius: 8px;
        padding: 14px;
        border-top: 4px solid;
    }
    .kpi-green { border-color: #10B981; }
    .kpi-amber { border-color: #F59E0B; }
    .kpi-red { border-color: #EF4444; }
    .kpi-neutral { border-color: #3B82F6; }
    .kpi-title { font-size: 13px; color: #94A3B8; margin-bottom: 2px; }
    .kpi-val { font-size: 24px; font-weight: 700; color: #F8FAFC; margin-bottom: 4px; }
    .kpi-sub { font-size: 12px; color: #CBD5E1; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🌱 Seed Blending & Inventory Allocation System</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automated Minimum-Labor Mixed-Integer Inventory Optimization Engine with Hard Bag-Unit Constraints</div>', unsafe_allow_html=True)

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

bag_size = st.sidebar.number_input(
    "Packaging Bag Unit Size (kg)",
    min_value=1,
    max_value=1000,
    value=25,
    step=5,
    help="Sets the integer decision variable unit size for the MILP solver. All allocations and constraints are solved in whole bag multiples."
)

st.sidebar.subheader("🎯 Quality Specification Limits")

max_ot2 = st.sidebar.slider(
    "Max Allowable OT-2 (Other Types) % Limit",
    min_value=0.10,
    max_value=5.00,
    value=1.50,
    step=0.05
)

min_gp = st.sidebar.slider(
    "Min Allowable GP (Germination Potential) % Limit",
    min_value=80.0,
    max_value=100.0,
    value=95.0,
    step=0.5
)

min_fc = st.sidebar.slider(
    "Min Allowable First Count (FC) % Limit",
    min_value=50.0,
    max_value=100.0,
    value=75.0,
    step=1.0
)

min_germ = st.sidebar.slider(
    "Min Allowable Final Germination % Limit",
    min_value=50.0,
    max_value=100.0,
    value=80.0,
    step=1.0
)

def run_milp_optimization(file_obj, target_qty, loc_option, max_ot2, min_gp, min_fc, min_germ, bag_size):
    if file_obj is None:
        try:
            file_obj = "Seed Blending Tool 2 (2).xlsx"
            xls = pd.ExcelFile(file_obj, engine='openpyxl')
        except Exception:
            return None, "Default dataset not found. Please upload an inventory Excel file via the sidebar.", None, None
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
        return None, "No inventory batches match the selected location filter.", None, None

    n = len(df_filtered)
    
    # Explicit integer bag target calculation: Ceiling round to ensure target demand is fully met
    bag_target = math.ceil(target_qty / bag_size)
    effective_target_kg = bag_target * bag_size
    
    # Max integer bags physically available per batch without violating stock
    max_bags_per_batch = np.floor(df_filtered['SAP Total'].values / bag_size)
    
    # Objective: Minimize total labor cost in bag decisions (y_i)
    c = df_filtered['Labor_Rank'].values * bag_size
    
    # Decision variables: y_i is integer count of bags for batch i
    integrality = np.ones(n)
    bounds = Bounds(lb=np.zeros(n), ub=max_bags_per_batch)
    
    # Linear Constraints on integer bag variables y_i:
    # 1. Total bags: sum(y_i) == bag_target
    # 2. OT-2% max: sum(y_i * (OT-2_i - max_ot2)) <= 0  (upper bound 0)
    # 3. GP% min:   sum(y_i * (min_gp - GP_i)) <= 0
    # 4. FC% min:   sum(y_i * (min_fc - FC_i)) <= 0
    # 5. Germ% min: sum(y_i * (min_germ - Germ_i)) <= 0
    
    A_rows = [
        np.ones(n),
        (df_filtered['OT-2%'].values - max_ot2),
        (min_gp - df_filtered['G.P %'].values),
        (min_fc - df_filtered['First Count'].values),
        (min_germ - df_filtered['Final G%'].values)
    ]
    
    lhs = [bag_target, -np.inf, -np.inf, -np.inf, -np.inf]
    rhs = [bag_target, 0.0, 0.0, 0.0, 0.0]
    
    constraints = LinearConstraint(A_rows, lhs, rhs)
    
    res = milp(c=c, integrality=integrality, bounds=bounds, constraints=constraints)
    
    if not res.success:
        total_stock = df_filtered['SAP Total'].sum()
        total_max_bags_kg = max_bags_per_batch.sum() * bag_size
        
        # Determine maximum achievable volume within all specs using MILP
        res_max_vol = milp(c=-np.ones(n) * bag_size, integrality=integrality, bounds=bounds,
                           constraints=LinearConstraint(A_rows[1:], [-np.inf]*4, [0.0]*4))
        max_blendable = -res_max_vol.fun if res_max_vol.success else 0.0

        # Assess single-parameter best achievable levels on whole bags
        eq_constr = LinearConstraint(np.ones(n), bag_target, bag_target)
        
        res_ot2 = milp(c=df_filtered['OT-2%'].values, integrality=integrality, bounds=bounds, constraints=eq_constr)
        best_ot2 = (np.sum(res_ot2.x * df_filtered['OT-2%'].values) / bag_target) if res_ot2.success else None

        res_gp = milp(c=-df_filtered['G.P %'].values, integrality=integrality, bounds=bounds, constraints=eq_constr)
        best_gp = (np.sum(res_gp.x * df_filtered['G.P %'].values) / bag_target) if res_gp.success else None

        res_fc = milp(c=-df_filtered['First Count'].values, integrality=integrality, bounds=bounds, constraints=eq_constr)
        best_fc = (np.sum(res_fc.x * df_filtered['First Count'].values) / bag_target) if res_fc.success else None

        res_germ = milp(c=-df_filtered['Final G%'].values, integrality=integrality, bounds=bounds, constraints=eq_constr)
        best_germ = (np.sum(res_germ.x * df_filtered['Final G%'].values) / bag_target) if res_germ.success else None

        violations = []
        if total_stock < effective_target_kg:
            violations.append(f"Total available stock ({total_stock:,.0f} kg) is less than target demand ({effective_target_kg:,.0f} kg).")
        elif total_max_bags_kg < effective_target_kg:
            violations.append(f"Total packageable stock in full {bag_size} kg bags ({total_max_bags_kg:,.0f} kg) cannot meet required {effective_target_kg:,.0f} kg.")
        
        if best_ot2 is not None and best_ot2 > max_ot2:
            violations.append(f"Cannot meet Max OT-2 {max_ot2:.2f}% — best achievable in full bag lots is {best_ot2:.2f}% (exceeds by {best_ot2 - max_ot2:.2f}pp).")
        if best_gp is not None and best_gp < min_gp:
            violations.append(f"Cannot meet Min GP {min_gp:.2f}% — best achievable in full bag lots is {best_gp:.2f}% (short by {min_gp - best_gp:.2f}pp).")
        if best_fc is not None and best_fc < min_fc:
            violations.append(f"Cannot meet Min First Count {min_fc:.2f}% — best achievable in full bag lots is {best_fc:.2f}% (short by {min_fc - best_fc:.2f}pp).")
        if best_germ is not None and best_germ < min_germ:
            violations.append(f"Cannot meet Min Germination {min_germ:.2f}% — best achievable in full bag lots is {best_germ:.2f}% (short by {min_germ - best_germ:.2f}pp).")

        failure_info = {
            'violations': violations,
            'max_blendable': max_blendable,
            'target_qty': effective_target_kg,
            'requested_qty': target_qty,
            'bag_size': bag_size
        }
        return None, None, None, failure_info

    # Solution Found
    df_res = df_filtered.copy()
    bag_counts = np.round(res.x).astype(int)
    df_res['Bags_Allocated'] = bag_counts
    df_res['Allocation Qty'] = bag_counts * bag_size
    
    df_allocated = df_res[df_res['Allocation Qty'] > 0].copy()
    
    # HARD ASSERTION: Verify no over-allocation against SAP Total
    over_allocated = df_allocated[df_allocated['Allocation Qty'] > df_allocated['SAP Total']]
    assert len(over_allocated) == 0, f"Over-allocation detected in batches: {over_allocated['Batch ID'].tolist()}"
    
    total_allocated = df_allocated['Allocation Qty'].sum()
    
    # Raw Solver Mathematical Products (for debug expander)
    df_allocated['ot2_product'] = df_allocated['Allocation Qty'] * df_allocated['OT-2%']
    df_allocated['gp_product'] = df_allocated['Allocation Qty'] * df_allocated['G.P %']
    df_allocated['fc_product'] = df_allocated['Allocation Qty'] * df_allocated['First Count']
    df_allocated['germ_product'] = df_allocated['Allocation Qty'] * df_allocated['Final G%']

    # True per-batch percentage specifications guarded against division by zero
    df_allocated['OT-2 %'] = np.where(df_allocated['Allocation Qty'] > 0, df_allocated['ot2_product'] / df_allocated['Allocation Qty'], 0.0)
    df_allocated['GP %'] = np.where(df_allocated['Allocation Qty'] > 0, df_allocated['gp_product'] / df_allocated['Allocation Qty'], 0.0)
    df_allocated['First Count %'] = np.where(df_allocated['Allocation Qty'] > 0, df_allocated['fc_product'] / df_allocated['Allocation Qty'], 0.0)
    df_allocated['Germination %'] = np.where(df_allocated['Allocation Qty'] > 0, df_allocated['germ_product'] / df_allocated['Allocation Qty'], 0.0)

    # Weighted Average Final Blended Quality Indicators
    blend_ot2 = df_allocated['ot2_product'].sum() / total_allocated
    blend_gp = df_allocated['gp_product'].sum() / total_allocated
    blend_fc = df_allocated['fc_product'].sum() / total_allocated
    blend_germ = df_allocated['germ_product'].sum() / total_allocated

    # HARD PASS/FAIL QUALITY VALIDATION
    pass_ot2 = (blend_ot2 <= max_ot2 + 1e-5)
    pass_gp = (blend_gp >= min_gp - 1e-5)
    pass_fc = (blend_fc >= min_fc - 1e-5)
    pass_germ = (blend_germ >= min_germ - 1e-5)
    all_specs_passed = pass_ot2 and pass_gp and pass_fc and pass_germ

    metrics = {
        'Target Needed': target_qty,
        'Effective Target': effective_target_kg,
        'Total Allocated': total_allocated,
        'Blended OT-2 %': blend_ot2,
        'Blended GP %': blend_gp,
        'Blended FC %': blend_fc,
        'Blended Germ %': blend_germ,
        'Limit OT-2': max_ot2,
        'Limit GP': min_gp,
        'Limit FC': min_fc,
        'Limit Germ': min_germ,
        'Bag Size': bag_size,
        'Passed All Specs': all_specs_passed,
        'Status': "✅ Priority 1 (Optimal MILP Minimum-Labor Blend Succeeded)" if all_specs_passed else "🚨 Quality Validation Failed on Rounded Solution"
    }

    display_cols = ['Batch ID', 'Location', 'Chamber-Bin', 'Labor_Rank', 'SAP Total', 'Allocation Qty', 'OT-2 %', 'GP %', 'First Count %', 'Germination %']
    debug_cols = ['Batch ID', 'Allocation Qty', 'ot2_product', 'gp_product', 'fc_product', 'germ_product']
    
    return metrics, df_allocated[display_cols].reset_index(drop=True), df_allocated[debug_cols].reset_index(drop=True), None

metrics, df_allocated, df_debug, failure_info = run_milp_optimization(
    uploaded_file, target_qty, location_option, max_ot2, min_gp, min_fc, min_germ, bag_size
)

if failure_info:
    st.error("🚨 **Optimization Infeasible: No MILP discrete bag configuration satisfies all active quality and stock constraints.**")
    st.markdown("#### Root Cause Breakdown:")
    for v in failure_info['violations']:
        st.markdown(f"- ❌ {v}")
    
    if failure_info['max_blendable'] > 0:
        st.info(f"💡 **Maximum blendable volume meeting all active specifications:** `{failure_info['max_blendable']:,.0f} kg` (out of `{failure_info['target_qty']:,.0f} kg` target in {failure_info['bag_size']} kg bags).")
    else:
        st.warning("💡 **No volume can be blended under the current strict quality constraints. Consider relaxing specification limits or broadening the location filter.**")

elif metrics:
    # Compute Headrooms
    headroom_ot2 = metrics['Limit OT-2'] - metrics['Blended OT-2 %']
    headroom_gp = metrics['Blended GP %'] - metrics['Limit GP']
    headroom_fc = metrics['Blended FC %'] - metrics['Limit FC']
    headroom_germ = metrics['Blended Germ %'] - metrics['Limit Germ']

    # Headroom Color Rules: negative = red (FAIL), 0-1pp = amber, >1pp = green
    def get_kpi_color(headroom):
        if headroom < -1e-5:
            return "kpi-red"
        elif headroom < 1.0:
            return "kpi-amber"
        else:
            return "kpi-green"

    color_ot2 = get_kpi_color(headroom_ot2)
    color_gp = get_kpi_color(headroom_gp)
    color_fc = get_kpi_color(headroom_fc)
    color_germ = get_kpi_color(headroom_germ)

    if not metrics['Passed All Specs']:
        st.error("🚨 **Quality Spec Violation Detected: The final allocation violated one or more target limits.**")
    else:
        st.success(metrics['Status'])

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""
        <div class="kpi-container kpi-neutral">
            <div class="kpi-title">Demand Target / Allocated</div>
            <div class="kpi-val">{metrics['Total Allocated']:,.0f} kg</div>
            <div class="kpi-sub">Target: {metrics['Target Needed']:,.0f} kg (in {metrics['Bag Size']} kg bags)</div>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-container {color_ot2}">
            <div class="kpi-title">Weighted OT-2 % (Other Types)</div>
            <div class="kpi-val">{metrics['Blended OT-2 %']:.2f}%</div>
            <div class="kpi-sub">Limit: {metrics['Limit OT-2']:.2f}% max | Headroom: {headroom_ot2:+.2f}pp</div>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="kpi-container {color_gp}">
            <div class="kpi-title">Weighted GP % (Germ. Potential)</div>
            <div class="kpi-val">{metrics['Blended GP %']:.2f}%</div>
            <div class="kpi-sub">Limit: {metrics['Limit GP']:.2f}% min | Headroom: {headroom_gp:+.2f}pp</div>
        </div>
        """, unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="kpi-container {color_fc}">
            <div class="kpi-title">Weighted First Count % (FC)</div>
            <div class="kpi-val">{metrics['Blended FC %']:.2f}%</div>
            <div class="kpi-sub">Limit: {metrics['Limit FC']:.2f}% min | Headroom: {headroom_fc:+.2f}pp</div>
        </div>
        """, unsafe_allow_html=True)
    with k5:
        st.markdown(f"""
        <div class="kpi-container {color_germ}">
            <div class="kpi-title">Weighted Germination %</div>
            <div class="kpi-val">{metrics['Blended Germ %']:.2f}%</div>
            <div class="kpi-sub">Limit: {metrics['Limit Germ']:.2f}% min | Headroom: {headroom_germ:+.2f}pp</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📋 Allocated Inventory Batches")

    df_table = df_allocated.copy()
    
    totals_row = {
        'Batch ID': 'TOTALS (Weighted Blend)',
        'Location': '—',
        'Chamber-Bin': '—',
        'Labor_Rank': '—',
        'SAP Total': df_table['SAP Total'].sum(),
        'Allocation Qty': df_table['Allocation Qty'].sum(),
        'OT-2 %': metrics['Blended OT-2 %'],
        'GP %': metrics['Blended GP %'],
        'First Count %': metrics['Blended FC %'],
        'Germination %': metrics['Blended Germ %']
    }
    
    df_display_with_total = pd.concat([df_table, pd.DataFrame([totals_row])], ignore_index=True)

    df_formatted = df_display_with_total.copy()
    df_formatted['SAP Total'] = df_formatted['SAP Total'].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int, float, np.number)) else str(x))
    df_formatted['Allocation Qty'] = df_formatted['Allocation Qty'].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int, float, np.number)) else str(x))
    df_formatted['OT-2 %'] = df_formatted['OT-2 %'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float, np.number)) else str(x))
    df_formatted['GP %'] = df_formatted['GP %'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float, np.number)) else str(x))
    df_formatted['First Count %'] = df_formatted['First Count %'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float, np.number)) else str(x))
    df_formatted['Germination %'] = df_formatted['Germination %'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float, np.number)) else str(x))

    st.dataframe(df_formatted, use_container_width=True, hide_index=True)
    
    diff_target = metrics['Effective Target'] - metrics['Target Needed']
    rounding_note = f"rounded up by {diff_target:,.0f} kg to the next whole bag lot" if diff_target > 0 else "exact whole bag match"
    st.caption(f"ℹ️ **MILP Bag Decision Rule:** Target demand ({metrics['Target Needed']:,.0f} kg) is satisfied in exact {metrics['Bag Size']} kg bag units ({metrics['Effective Target']:,.0f} kg total, {rounding_note}). Every batch allocation is guaranteed $\\le$ SAP Total. *Note: Blended lot requires a fresh laboratory purity and germination certificate before commercial packaging and labelling. Displayed quality figures are computed projections.*")

    # Corrected Raw String Rendering for LaTeX
    with st.expander("🔍 Solver detail (debug)"):
        st.markdown(r"Raw mathematical product values ($\text{Allocation Qty} \times \text{Spec}$) used internally by the linear optimization solver:")
        st.dataframe(df_debug, use_container_width=True, hide_index=True)

    def create_result_excel(df_out, metrics_dict, totals):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            summary_df = pd.DataFrame([
                ["Target Demand Needed (kg)", metrics_dict['Target Needed']],
                ["Effective Demand Allocated (kg)", metrics_dict['Total Allocated']],
                ["Packaging Bag Unit Size (kg)", metrics_dict['Bag Size']],
                ["Weighted OT-2 (Other Types) %", metrics_dict['Blended OT-2 %']],
                ["Weighted GP (Germination Potential) %", metrics_dict['Blended GP %']],
                ["Weighted First Count (FC) %", metrics_dict['Blended FC %']],
                ["Weighted Final Germination %", metrics_dict['Blended Germ %']],
                ["Optimization Status", metrics_dict['Status']]
            ], columns=["Metric", "Value"])
            summary_df.to_excel(writer, sheet_name="Executive_Summary", index=False)
            
            export_df = pd.concat([df_out, pd.DataFrame([totals])], ignore_index=True)
            export_df.to_excel(writer, sheet_name="Allocated_Batches", index=False)
        return buffer.getvalue()

    st.subheader("📤 Export Reports")
    c1, c2 = st.columns(2)
    with c1:
        try:
            excel_report_bytes = create_result_excel(df_allocated, metrics, totals_row)
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
        csv_data = pd.concat([df_allocated, pd.DataFrame([totals_row])], ignore_index=True).to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Download Allocation CSV",
            data=csv_data,
            file_name="Seed_Allocation_Batches.csv",
            mime="text/csv",
            key="btn_csv_dl"
        )