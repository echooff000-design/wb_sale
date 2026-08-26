import streamlit as st
import pandas as pd
import requests
import io
import datetime
import re

# Set page layout to wide and dark-friendly
st.set_page_config(page_title="WB Sale Data", layout="wide", initial_sidebar_state="expanded")

# --- CUSTOM CSS FOR STYLING & MOBILE RESPONSIVENESS ---
st.markdown("""
<style>
    .block-container { padding: 1rem 1rem 2rem 1rem; max-width: 100%; }
    .metric-box { background: #1e293b; padding: 12px; border-radius: 8px; border: 1px solid #334155; margin-bottom: 8px; }
    .custom-table { width: 100%; border-collapse: collapse; font-family: Calibri, 'Segoe UI', sans-serif; font-size: 13px; color: #000; background: #fff; margin-bottom: 1rem; }
    .custom-table th, .custom-table td { border: 1px solid #d3d3d3; padding: 5px 8px; text-align: center; white-space: nowrap; }
    .custom-table th { background-color: #D9E1F2; font-weight: bold; }
    .custom-table th:first-child, .custom-table td:first-child { text-align: left; position: sticky; left: 0; background-color: #F2F2F2; z-index: 2; }
    .custom-table th:first-child { background-color: #D9E1F2; z-index: 3; }
    .subtotal-row { font-weight: bold; background-color: #F2F2F2; text-align: left; }
    .grand-total-row { background-color: #D9E1F2; font-weight: bold; }
    .hl-green { background-color: #def7ec !important; color: #03543f !important; font-weight: bold; }
    .hl-red { background-color: #fde8e8 !important; color: #9b1c1c !important; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- CONFIGURATION & MASTER STRUCTURE ---
RAW_URL = "https://tilaknagarindustries-my.sharepoint.com/:x:/g/personal/andebnath_tilind_com/IQBQOO7YAaZLQIlMbx55OXr1AdvPSbB0XsQEEmRvQMvFdBY?download=1"

MASTER_STRUCTURE = [
    {"seg": "Deluxe-Whisky", "brands": ["IBDC", "N1WSUP", "OCBL", "GGSW", "Green Label", "IQ", "MCD Lux", "Mountain Oak"]},
    {"seg": "Semi Premium-Whisky", "brands": ["MHW", "All Season", "Brothers", "GRAYSON'S Maxx", "OakInt", "RCW", "RGW", "ROCKFORD", "RSBS", "RSDD", "RSW", "SRB7", "Whiskots", "GRR"]},
    {"seg": "Deluxe-Gin", "brands": ["BLGLM", "BLGOR", "Big Ben", "Blue Riband"]},
    {"seg": "Premium-Brandy", "brands": ["Monarch"]},
    {"seg": "Premium-Gin", "brands": ["SMG", "SMGP"]},
    {"seg": "Semi Premium-Brandy", "brands": ["MHFB"]},
    {"seg": "Single Malt-Scotch", "brands": ["SIW"]}
]
MARKED_BRANDS = ['IBDC', 'MHW', 'BLGLM', 'BLGOR', 'Monarch', 'SMG', 'SMGP', 'MHFB', 'SIW']

# --- DATA FETCHING & PARSING ENGINE ---
@st.cache_data(ttl=1800, show_spinner=False)
def load_dataset_from_sharepoint():
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    response = requests.get(RAW_URL, headers=headers, timeout=60)
    response.raise_for_status()
    
    excel_bytes = io.BytesIO(response.content)
    xls = pd.ExcelFile(excel_bytes)
    
    this_m = pd.read_excel(xls, "This Month") if "This Month" in xls.sheet_names else pd.DataFrame()
    last_m = pd.read_excel(xls, "Last Month") if "Last Month" in xls.sheet_names else pd.DataFrame()
    tgt_df = pd.read_excel(xls, "Target Data") if "Target Data" in xls.sheet_names else pd.DataFrame()
    out_master = pd.read_excel(xls, "Outlet Master") if "Outlet Master" in xls.sheet_names else pd.DataFrame()
    users_df = pd.read_excel(xls, "Users") if "Users" in xls.sheet_names else pd.DataFrame()
    
    # Extract Last Synced Date and Elapsed Days
    sync_date_str = "Live"
    days_elapsed = 21
    if not users_df.empty:
        for col in ["Date", "date", users_df.columns[min(5, len(users_df.columns)-1)]]:
            if col in users_df.columns and pd.notna(users_df[col].iloc[0]):
                sync_date_str = str(users_df[col].iloc[0]).strip()
                match = re.search(r'\b(\d{1,2})\b', sync_date_str)
                if match:
                    days_elapsed = int(match.group(1))
                break

    # Build Outlet Metadata Map
    outlet_meta = {}
    if not out_master.empty:
        for _, r in out_master.iterrows():
            lic = str(r.get("LIC No", r.get("Outlet Name", ""))).strip()
            if lic:
                outlet_meta[lic] = {
                    "group": str(r.get("Group", "Unassigned")).strip(),
                    "zone": str(r.get("Zone", "West Bengal")).strip(),
                    "asm": str(r.get("ASM", "Unassigned")).strip(),
                    "tse": str(r.get("TSE", "Unassigned")).strip()
                }

    # Normalize & Aggregate Sales Records
    combined = {}
    def process_sheet(df, metric_key):
        if df.empty:
            return
        vol_col = "Volume" if "Volume" in df.columns else ("Value" if "Value" in df.columns else df.columns[-1])
        for _, r in df.iterrows():
            lic = str(r.get("LIC No", r.get("Outlet Name", ""))).strip()
            if not lic:
                continue
            brand = str(r.get("Brand", "")).strip()
            if brand == "IBW":
                brand = "IBDC"
            seg = str(r.get("Segment", "Deluxe-Whisky")).strip()
            if seg == "Deluxe Plus-Whisky":
                seg = "Deluxe-Whisky"
            
            val = float(r.get(vol_col, 0) or 0)
            uid = f"{lic}_{brand}"
            if uid not in combined:
                meta = outlet_meta.get(lic, {
                    "group": "Unassigned", 
                    "zone": "West Bengal", 
                    "asm": str(r.get("ASM", "Unassigned")).strip(), 
                    "tse": str(r.get("TSE", "Unassigned")).strip()
                })
                combined[uid] = {
                    "lic": lic,
                    "outlet": str(r.get("Outlet Name", lic)).strip(),
                    "group": meta["group"],
                    "zone": meta["zone"],
                    "asm": meta["asm"],
                    "tse": meta["tse"],
                    "seg": seg,
                    "brand": brand,
                    "tm": 0.0, "lm": 0.0, "tgt": 0.0
                }
            combined[uid][metric_key] += val

    process_sheet(this_m, "tm")
    process_sheet(last_m, "lm")
    process_sheet(tgt_df, "tgt")

    return pd.DataFrame(list(combined.values())), users_df, sync_date_str, days_elapsed

# --- LOAD DATA ---
with st.spinner("Connecting to live sales data..."):
    try:
        sales_df, users_df, sync_date_str, days_elapsed = load_dataset_from_sharepoint()
    except Exception as e:
        st.error(f"Failed to load data from live source: {e}")
        st.stop()

# --- TOP HEADER BAR ---
col_head1, col_head2 = st.columns([3, 1])
with col_head1:
    st.subheader("WB Sale Data Dashboard")
    st.caption(f"🕒 **Last Sync Date:** {sync_date_str} | Days Elapsed: {days_elapsed}")
with col_head2:
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# --- SIDEBAR CASCADING FILTERS ---
st.sidebar.header("🔍 Filters")

all_groups = sorted([g for g in sales_df["group"].dropna().unique() if g != "Unassigned"])
sel_group = st.sidebar.selectbox("Group Filter", ["All"] + all_groups)

f_df = sales_df if sel_group == "All" else sales_df[sales_df["group"] == sel_group]
all_asms = sorted([a for a in f_df["asm"].dropna().unique() if a != "Unassigned"])
sel_asm = st.sidebar.selectbox("ASM Filter", ["All"] + all_asms)

if sel_asm != "All":
    f_df = f_df[f_df["asm"] == sel_asm]
all_tses = sorted([t for t in f_df["tse"].dropna().unique() if t != "Unassigned"])
sel_tse = st.sidebar.selectbox("TSE Filter", ["All"] + all_tses)

if sel_tse != "All":
    f_df = f_df[f_df["tse"] == sel_tse]
all_lics = sorted([l for l in f_df["lic"].dropna().unique() if l])
sel_lic = st.sidebar.selectbox("LIC No Filter", ["All"] + all_lics)

if sel_lic != "All":
    f_df = f_df[f_df["lic"] == sel_lic]
all_outlets = sorted([o for o in f_df["outlet"].dropna().unique() if o])
sel_outlet = st.sidebar.selectbox("Outlet Filter", ["All"] + all_outlets)

search_term = st.sidebar.text_input("Instant Search (Outlet Name / LIC):").strip().lower()

# Apply final filtering
filtered_df = f_df if sel_outlet == "All" else f_df[f_df["outlet"] == sel_outlet]
if search_term:
    filtered_df = filtered_df[
        filtered_df["outlet"].str.lower().str.contains(search_term, na=False) |
        filtered_df["lic"].str.lower().str.contains(search_term, na=False)
    ]

# --- MAIN NAVIGATION TABS ---
tab_vol, tab_ms, tab_dash, tab_ask = st.tabs(["📦 Volume", "📈 Ms%", "📊 Dashboard", "💬 Ask Assistant"])

# ================= TAB 1: VOLUME MATRIX =================
with tab_vol:
    html_rows = ""
    gt_lm, gt_tgt, gt_tm, gt_bal = 0, 0, 0, 0
    for g in MASTER_STRUCTURE:
        seg_data = filtered_df[filtered_df["seg"] == g["seg"]]
        s_lm = seg_data["lm"].sum()
        s_tgt = seg_data["tgt"].sum()
        s_tm = seg_data["tm"].sum()
        html_rows += f'<tr class="subtotal-row"><td>{g["seg"]}</td><td>{int(round(s_lm)):,}</td><td>{int(round(s_tgt)):,}</td><td>{int(round(s_tm)):,}</td><td></td></tr>'
        for b in g["brands"]:
            b_data = seg_data[seg_data["brand"] == b]
            lm = b_data["lm"].sum()
            tgt = b_data["tgt"].sum()
            tm = b_data["tm"].sum()
            is_m = b in MARKED_BRANDS
            bal = (tgt - tm) if is_m else ""
            if is_m:
                gt_bal += (tgt - tm)
            hl = ("hl-red" if tm < tgt else "hl-green") if is_m else ""
            brand_style = 'style="text-align:left; padding-left:14px; background:#EBF5FB; font-weight:bold;"' if is_m else 'style="text-align:left; padding-left:14px;"'
            bal_str = f"{int(round(bal)):,}" if bal != "" else ""
            html_rows += f'<tr><td {brand_style}>{b}</td><td>{int(round(lm)):,}</td><td>{int(round(tgt)):,}</td><td class="{hl}">{int(round(tm)):,}</td><td class="{hl}">{bal_str}</td></tr>'
        gt_lm += s_lm
        gt_tgt += s_tgt
        gt_tm += s_tm

    html_table = f"""
    <div style="overflow-x:auto;">
    <table class="custom-table">
        <thead>
            <tr><th>Brand</th><th>LM</th><th>TGT</th><th>TM</th><th>BAL</th></tr>
        </thead>
        <tbody>
            {html_rows}
            <tr class="grand-total-row"><td>Grand Total</td><td>{int(round(gt_lm)):,}</td><td>{int(round(gt_tgt)):,}</td><td>{int(round(gt_tm)):,}</td><td>{int(round(gt_bal)):,}</td></tr>
        </tbody>
    </table>
    </div>
    """
    st.markdown(html_table, unsafe_allow_html=True)

# ================= TAB 2: MS% MATRIX =================
with tab_ms:
    gt_lm = filtered_df["lm"].sum() or 1.0
    gt_tm = filtered_df["tm"].sum() or 1.0
    html_ms = ""
    for g in MASTER_STRUCTURE:
        seg_data = filtered_df[filtered_df["seg"] == g["seg"]]
        s_lm = seg_data["lm"].sum()
        s_tm = seg_data["tm"].sum()
        s_lm_p = (s_lm / gt_lm) * 100
        s_tm_p = (s_tm / gt_tm) * 100
        s_diff = s_tm_p - s_lm_p
        html_ms += f'<tr class="subtotal-row"><td>{g["seg"]}</td><td>{s_lm_p:.1f}%</td><td>{s_tm_p:.1f}%</td><td>{s_diff:+.1f}%</td></tr>'
        for b in g["brands"]:
            b_data = seg_data[seg_data["brand"] == b]
            lm = b_data["lm"].sum()
            tm = b_data["tm"].sum()
            b_lm_p = (lm / s_lm * 100) if s_lm > 0 else 0.0
            b_tm_p = (tm / s_tm * 100) if s_tm > 0 else 0.0
            grw = b_tm_p - b_lm_p
            hl = "hl-green" if grw > 0 else ("hl-red" if grw < 0 else "")
            html_ms += f'<tr><td style="text-align:left; padding-left:14px;">{b}</td><td>{b_lm_p:.1f}%</td><td>{b_tm_p:.1f}%</td><td class="{hl}">{grw:+.1f}%</td></tr>'

    html_ms_table = f"""
    <div style="overflow-x:auto;">
    <table class="custom-table">
        <thead>
            <tr><th>Brand</th><th>LM</th><th>TM</th><th>GRW</th></tr>
        </thead>
        <tbody>
            {html_ms}
            <tr class="grand-total-row"><td>Grand Total</td><td>100.0%</td><td>100.0%</td><td></td></tr>
        </tbody>
    </table>
    </div>
    """
    st.markdown(html_ms_table, unsafe_allow_html=True)

# ================= TAB 3: HIERARCHY DASHBOARDS =================
with tab_dash:
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["Target vs Ach", "MS% Details", "WOD Details"])
    
    def calc_ms(df_sub, brand):
        segs = ["Semi Premium-Whisky"] if brand == "MHW" else ["Deluxe-Whisky", "Deluxe Plus-Whisky"]
        btm = df_sub[df_sub["brand"] == brand]["tm"].sum()
        dtm = df_sub[df_sub["seg"].isin(segs)]["tm"].sum()
        return (btm / dtm * 100) if dtm > 0 else 0.0

    def calc_lm_ms(df_sub, brand):
        segs = ["Semi Premium-Whisky"] if brand == "MHW" else ["Deluxe-Whisky", "Deluxe Plus-Whisky"]
        blm = df_sub[df_sub["brand"] == brand]["lm"].sum()
        dlm = df_sub[df_sub["seg"].isin(segs)]["lm"].sum()
        return (blm / dlm * 100) if dlm > 0 else 0.0

    # Sub-tab 1: Target vs Achieved
    with sub_tab1:
        def build_h1_row(name, sub, cls, pad):
            i_lm = sub[sub["brand"] == "IBDC"]["lm"].sum()
            i_tgt = sub[sub["brand"] == "IBDC"]["tgt"].sum()
            i_tm = sub[sub["brand"] == "IBDC"]["tm"].sum()
            m_lm = sub[sub["brand"] == "MHW"]["lm"].sum()
            m_tgt = sub[sub["brand"] == "MHW"]["tgt"].sum()
            m_tm = sub[sub["brand"] == "MHW"]["tm"].sum()
            return f'<tr class="{cls}"><td style="text-align:left; padding-left:{pad}px;">{name}</td><td>{int(round(i_lm)):,}</td><td>{int(round(i_tgt)):,}</td><td>{int(round(i_tm)):,}</td><td>{calc_ms(sub, "IBDC"):.1f}%</td><td>{int(round(m_lm)):,}</td><td>{int(round(m_tgt)):,}</td><td>{int(round(m_tm)):,}</td><td>{calc_ms(sub, "MHW"):.1f}%</td></tr>'

        h1_body = build_h1_row("West Bengal", filtered_df, "grand-total-row", 6)
        for z in sorted([x for x in filtered_df["zone"].dropna().unique() if x]):
            z_df = filtered_df[filtered_df["zone"] == z]
            h1_body += build_h1_row(z, z_df, "subtotal-row", 6)
            for a in sorted([x for x in z_df["asm"].dropna().unique() if x]):
                a_df = z_df[z_df["asm"] == a]
                h1_body += build_h1_row(a, a_df, "subtotal-row", 14)
                for t in sorted([x for x in a_df["tse"].dropna().unique() if x]):
                    t_df = a_df[a_df["tse"] == t]
                    h1_body += build_h1_row(t, t_df, "", 22)

        st.markdown(f"""
        <div style="overflow-x:auto;">
        <table class="custom-table">
            <thead>
                <tr><th rowspan="2">ZONE/ASM/TSE</th><th colspan="4">IBDC</th><th colspan="4">MHW</th></tr>
                <tr><th>LM</th><th>Target</th><th>MTD</th><th>MS%</th><th>LM</th><th>Target</th><th>MTD</th><th>MS%</th></tr>
            </thead>
            <tbody>{h1_body}</tbody>
        </table>
        </div>
        """, unsafe_allow_html=True)

    # Sub-tab 2: MS% Details
    with sub_tab2:
        h2_brands = ["IBDC", "MCD Lux", "IQ", "N1WSUP", "OCBL", "RSW", "SRB7", "RGW", "MHW"]
        def build_h2_row(name, sub, cls, pad):
            cols = ""
            for b in h2_brands:
                lm_ms = calc_lm_ms(sub, b)
                tm_ms = calc_ms(sub, b)
                diff = tm_ms - lm_ms
                color = "#03543f" if diff >= 0 else "#9b1c1c"
                cols += f"<td>{lm_ms:.1f}%</td><td>{tm_ms:.1f}%</td><td style='color:{color}; font-weight:bold;'>{diff:+.1f}%</td>"
            return f'<tr class="{cls}"><td style="text-align:left; padding-left:{pad}px;">{name}</td>{cols}</tr>'

        h2_body = build_h2_row("West Bengal", filtered_df, "grand-total-row", 6)
        for z in sorted([x for x in filtered_df["zone"].dropna().unique() if x]):
            z_df = filtered_df[filtered_df["zone"] == z]
            h2_body += build_h2_row(z, z_df, "subtotal-row", 6)
            for a in sorted([x for x in z_df["asm"].dropna().unique() if x]):
                a_df = z_df[z_df["asm"] == a]
                h2_body += build_h2_row(a, a_df, "subtotal-row", 14)
                for t in sorted([x for x in a_df["tse"].dropna().unique() if x]):
                    t_df = a_df[a_df["tse"] == t]
                    h2_body += build_h2_row(t, t_df, "", 22)

        header_b = "".join([f'<th colspan="3">{b}</th>' for b in h2_brands])
        header_sub = "".join(['<th>LM</th><th>MTD</th><th>diff</th>' for _ in h2_brands])
        st.markdown(f"""
        <div style="overflow-x:auto;">
        <table class="custom-table">
            <thead>
                <tr><th rowspan="2">ZONE/ASM/TSE</th>{header_b}</tr>
                <tr>{header_sub}</tr>
            </thead>
            <tbody>{h2_body}</tbody>
        </table>
        </div>
        """, unsafe_allow_html=True)

    # Sub-tab 3: WOD Details
    with sub_tab3:
        h3_brands = ["IBDC", "MCD Lux", "IQ", "MHW"]
        def build_h3_row(name, sub, cls, pad):
            cols = ""
            for b in h3_brands:
                lm_u = sub[(sub["brand"] == b) & (sub["lm"] > 0)]["lic"].nunique()
                tm_u = sub[(sub["brand"] == b) & (sub["tm"] > 0)]["lic"].nunique()
                diff = tm_u - lm_u
                color = "#03543f" if diff >= 0 else "#9b1c1c"
                cols += f"<td>{lm_u:,}</td><td>{tm_u:,}</td><td style='color:{color}; font-weight:bold;'>{diff:+d}</td>"
            return f'<tr class="{cls}"><td style="text-align:left; padding-left:{pad}px;">{name}</td>{cols}</tr>'

        h3_body = build_h3_row("West Bengal", filtered_df, "grand-total-row", 6)
        for z in sorted([x for x in filtered_df["zone"].dropna().unique() if x]):
            z_df = filtered_df[filtered_df["zone"] == z]
            h3_body += build_h3_row(z, z_df, "subtotal-row", 6)
            for a in sorted([x for x in z_df["asm"].dropna().unique() if x]):
                a_df = z_df[z_df["asm"] == a]
                h3_body += build_h3_row(a, a_df, "subtotal-row", 14)
                for t in sorted([x for x in a_df["tse"].dropna().unique() if x]):
                    t_df = a_df[a_df["tse"] == t]
                    h3_body += build_h3_row(t, t_df, "", 22)

        header3_b = "".join([f'<th colspan="3">{b}</th>' for b in h3_brands])
        header3_sub = "".join(['<th>LM</th><th>MTD</th><th>diff</th>' for _ in h3_brands])
        st.markdown(f"""
        <div style="overflow-x:auto;">
        <table class="custom-table">
            <thead>
                <tr><th rowspan="2">Unique Billing Outlet<br>ZONE/ASM/TSE</th>{header3_b}</tr>
                <tr>{header3_sub}</tr>
            </thead>
            <tbody>{h3_body}</tbody>
        </table>
        </div>
        """, unsafe_allow_html=True)

# ================= TAB 4: ASK ASSISTANT QUERY ENGINE =================
with tab_ask:
    query_type = st.selectbox(
        "Choose Query Analysis:",
        [
            "TIL Non Billed Outlets",
            "Deluxe Industry >= 30 CS but IBDC Not Billed",
            "Semi Premium Whisky Industry >= 50 CS but MHW Not Billed",
            "Magic Moments Billed but BLG Not Billed",
            "MCD Lux Billed but IBDC Not Billed",
            "IQ Billed but IBDC Not Billed",
            "RSW Billed but MHW Not Billed",
            "RGW Billed but MHW Not Billed",
            "SRB7 Billed but MHW Not Billed",
            "RCW Billed but MHW Not Billed",
            "All Season Billed but MHW Not Billed",
            "Brand-wise L3M Daily Run vs Current Month Daily Run"
        ]
    )

    if query_type == "Brand-wise L3M Daily Run vs Current Month Daily Run":
        daily_rows = []
        html_daily = ""
        gt_l3m, gt_tm = 0, 0
        for g in MASTER_STRUCTURE:
            s_df = filtered_df[filtered_df["seg"] == g["seg"]]
            s_l3m = s_df["lm"].sum()
            s_tm = s_df["tm"].sum()
            s_l3m_d = s_l3m / 90.0
            s_tm_d = s_tm / float(days_elapsed)
            s_grw = s_tm_d - s_l3m_d
            s_grw_p = (s_grw / s_l3m_d * 100) if s_l3m_d > 0 else 0.0
            html_daily += f'<tr class="subtotal-row"><td>{g["seg"]}</td><td>{int(round(s_l3m)):,}</td><td>{s_l3m_d:.1f}</td><td>{int(round(s_tm)):,}</td><td>{s_tm_d:.1f}</td><td>{s_grw:+.1f}</td><td>{s_grw_p:+.1f}%</td></tr>'
            for b in g["brands"]:
                b_df = s_df[s_df["brand"] == b]
                l3m = b_df["lm"].sum()
                tm = b_df["tm"].sum()
                l3m_d = l3m / 90.0
                tm_d = tm / float(days_elapsed)
                grw = tm_d - l3m_d
                grw_p = (grw / l3m_d * 100) if l3m_d > 0 else 0.0
                hl = "hl-green" if grw > 0 else ("hl-red" if grw < 0 else "")
                html_daily += f'<tr><td style="text-align:left; padding-left:14px;">{b}</td><td>{int(round(l3m)):,}</td><td>{l3m_d:.1f}</td><td>{int(round(tm)):,}</td><td>{tm_d:.1f}</td><td class="{hl}">{grw:+.1f}</td><td class="{hl}">{grw_p:+.1f}%</td></tr>'
                daily_rows.append({"Segment": g["seg"], "Brand": b, "L3M Total": int(round(l3m)), "L3M Daily": round(l3m_d, 1), "TM Total": int(round(tm)), "TM Daily": round(tm_d, 1), "Growth CS": round(grw, 1), "Growth %": f"{grw_p:.1f}%"})
            gt_l3m += s_l3m
            gt_tm += s_tm

        gt_l3m_d = gt_l3m / 90.0
        gt_tm_d = gt_tm / float(days_elapsed)
        gt_grw = gt_tm_d - gt_l3m_d
        gt_grw_p = (gt_grw / gt_l3m_d * 100) if gt_l3m_d > 0 else 0.0

        st.markdown(f"""
        <div style="overflow-x:auto;">
        <table class="custom-table">
            <thead>
                <tr><th>Brand</th><th>L3M Total</th><th>L3M Daily (/90)</th><th>TM Total</th><th>TM Daily (/{days_elapsed}D)</th><th>Growth (CS)</th><th>Growth %</th></tr>
            </thead>
            <tbody>
                {html_daily}
                <tr class="grand-total-row"><td>Grand Total</td><td>{int(round(gt_l3m)):,}</td><td>{gt_l3m_d:.1f}</td><td>{int(round(gt_tm)):,}</td><td>{gt_tm_d:.1f}</td><td>{gt_grw:+.1f}</td><td>{gt_grw_p:+.1f}%</td></tr>
            </tbody>
        </table>
        </div>
        """, unsafe_allow_html=True)
        
        # Download Button
        res_df = pd.DataFrame(daily_rows)
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine='xlsxwriter') as writer:
            res_df.to_excel(writer, index=False, sheet_name="DailyRun")
        st.download_button("📥 Download Excel Report", data=out_buf.getvalue(), file_name="Daily_Run_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        # Outlet-Level Gap Filtering
        gap_records = []
        for out, grp in filtered_df.groupby("outlet"):
            lic = grp["lic"].iloc[0]
            asm = grp["asm"].iloc[0]
            tse = grp["tse"].iloc[0]
            matched = False
            vol = 0.0

            if "Deluxe Industry >=" in query_type:
                d_vol = grp[grp["seg"].str.contains("Deluxe", na=False)]["lm"].sum()
                i_vol = grp[grp["brand"] == "IBDC"]["tm"].sum()
                if d_vol >= 30 and i_vol == 0:
                    matched, vol = True, d_vol
            elif "Semi Premium Whisky Industry >=" in query_type:
                sp_vol = grp[grp["seg"].str.contains("Semi Premium-Whisky", na=False)]["lm"].sum()
                m_vol = grp[grp["brand"] == "MHW"]["tm"].sum()
                if sp_vol >= 50 and m_vol == 0:
                    matched, vol = True, sp_vol
            elif "TIL Non Billed Outlets" in query_type:
                b_vol = grp["lm"].sum()
                t_vol = grp[grp["brand"].isin(MARKED_BRANDS)]["tm"].sum()
                if b_vol > 0 and t_vol == 0:
                    matched, vol = True, b_vol
            elif "Billed but" in query_type:
                driver, target = "MCD Lux", "IBDC"
                if "Magic Moments" in query_type: driver, target = "Big Ben", "BLGLM"
                elif "IQ" in query_type: driver, target = "IQ", "IBDC"
                elif "RSW" in query_type: driver, target = "RSW", "MHW"
                elif "RGW" in query_type: driver, target = "RGW", "MHW"
                elif "SRB7" in query_type: driver, target = "SRB7", "MHW"
                elif "RCW" in query_type: driver, target = "RCW", "MHW"
                elif "All Season" in query_type: driver, target = "All Season", "MHW"
                
                drv_v = grp[grp["brand"] == driver]["lm"].sum()
                tgt_v = grp[grp["brand"] == target]["tm"].sum()
                if drv_v > 0 and tgt_v == 0:
                    matched, vol = True, drv_v

            if matched:
                gap_records.append({"LIC No": lic, "Outlet Name": out, "ASM": asm, "TSE": tse, "Volume (CS)": int(round(vol))})

        gap_df = pd.DataFrame(gap_records)
        st.info(f"Total Outlets Found: {len(gap_df):,}")
        if not gap_df.empty:
            st.dataframe(gap_df, use_container_width=True, hide_index=True)
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine='xlsxwriter') as writer:
                gap_df.to_excel(writer, index=False, sheet_name="GapAnalysis")
            st.download_button("📥 Download Excel Report", data=out_buf.getvalue(), file_name=f"{query_type.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("🎉 No gap outlets found for the selected filter criteria!")
