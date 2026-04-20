"""
儲互社分析系統 — 2026 官方正式部署版 (UI & 版面極致優化版)
"""

import io
import uuid
import logging
import traceback

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from supabase import create_client, Client

# ──────────────────────────────────────────────
# 🛠️ 全域配置
# ──────────────────────────────────────────────
CONFIG = {
    "BUCKET_NAME":  st.secrets.get("BUCKET_NAME", "excel-reports"),
    "APP_BASE_URL": "https://8asdxeziyl2ozfrmkpzof3.streamlit.app", 
    "MAX_ATTEMPTS": 5,
    "THEME_BG": "#F0F4F8", # 網頁底色
    "SHEETS": {
        "MAIN":   "社務及資金運用情形",
        "LOAN":   "放款及逾期放款",
        "REGION": "區域分類表"
    },
    "THRESHOLDS": {
        "high_risk_ovd":    st.secrets.get("thresholds", {}).get("high_risk_ovd", 0.1),
        "liquidity_loan":   st.secrets.get("thresholds", {}).get("liquidity_loan", 0.9),
        "idle_loan":        st.secrets.get("thresholds", {}).get("idle_loan", 0.3),
        "stable_loan_min":  st.secrets.get("thresholds", {}).get("stable_loan_min", 0.4),
        "stable_loan_max":  st.secrets.get("thresholds", {}).get("stable_loan_max", 0.8),
        "ovd_safe_line":    st.secrets.get("thresholds", {}).get("ovd_safe_line", 0.02),
    }
}

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# 🎨 頁面與樣式設定
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="儲互社分析系統",
    layout="wide", # 預設寬螢幕
    page_icon="🏦",
    initial_sidebar_state="collapsed",
)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600;700&display=swap');
html, body, [data-testid="stAppViewContainer"] {{ 
    font-family: 'Noto Sans TC', sans-serif !important; 
    background-color: {CONFIG['THEME_BG']} !important; 
    color: #1A202C !important; 
}}

/* 強制填滿版面：當側邊欄收起時，右側內容自動擴張至 100% */
[data-testid="stMainBlockContainer"] {{
    max-width: 100% !important;
    padding-top: 4rem !important; /* 增加頂部間距，確保標題不貼邊 */
    padding-left: 1.5rem !important;
    padding-right: 1.5rem !important;
    padding-bottom: 1.5rem !important;
}}

/* 標題樣式優化 */
.responsive-h1 {{ 
    font-size: 2.25rem; 
    font-weight: 700; 
    margin-bottom: 2rem !important; 
    color: #1E293B;
}}
.responsive-h2 {{ 
    font-size: 1.75rem; 
    font-weight: 700; 
    margin-bottom: 1.5rem !important; 
    color: #1E293B;
}}

@media (max-width: 640px) {{
    [data-testid="stMainBlockContainer"] {{ 
        padding-left: 1rem !important; 
        padding-right: 1rem !important; 
        padding-top: 4rem !important; /* 增加頂部間距，避免被遮蓋 */
    }}
    .stat-card {{ min-height: auto !important; }}
    .responsive-h1 {{ font-size: 1.5rem !important; }} /* 手機版主標題縮小 */
    .responsive-h2 {{ font-size: 1.25rem !important; }} /* 手機版登入標題縮小 */
}}

/* 側邊欄外觀 */
[data-testid="stSidebar"] {{ background-color: #1E293B !important; }}
[data-testid="stSidebar"] * {{ color: #E2E8F0 !important; }}
[data-testid="stSidebar"] hr {{ border-color: #334155 !important; margin: 1.5rem 0 !important; }}
[data-testid="stSidebar"] .stButton > button {{
    background: #334155; color: #E2E8F0 !important; border: 1px solid #475569; border-radius: 10px;
    padding: 0.5rem 1rem; font-weight: 600; width: 100%; transition: all 0.2s;
}}
[data-testid="stSidebar"] .stButton > button:hover {{ background: #475569; border-color: #64748B; transform: translateY(-1px); }}

/* 連結框 */
.stCodeBlock {{ border-radius: 10px !important; background: #0F172A !important; border: 1px solid #334155 !important; }}

[data-testid="stVerticalBlockBorderWrapper"] {{ border-radius: 20px !important; background: white !important; padding: 1.5rem !important; box-shadow: 0 10px 25px rgba(0,0,0,0.05) !important; }}
.stat-card {{ background: white; border-radius: 14px; border: 1px solid #E2E8F0; margin-bottom: 1rem; min-height: 180px; display: flex; flex-direction: column; overflow: hidden; }}
.card-header {{ padding: 10px; color: #FFF !important; font-weight: 700; text-align: center; }}
.hdr-red {{ background: linear-gradient(135deg, #EF4444, #991B1B); }}
.hdr-orange {{ background: linear-gradient(135deg, #F59E0B, #92400E); }}
.hdr-blue {{ background: linear-gradient(135deg, #3B82F6, #1E40AF); }}
.hdr-green {{ background: linear-gradient(135deg, #10B981, #065F46); }}
.name-tag {{ display: inline-block; background: #F1F5F9; color: #1A202C !important; padding: 3px 10px; border-radius: 8px; margin: 3px; font-size: .82rem; border: 1px solid #CBD5E1; font-weight: 600; }}
.badge-admin {{ background: #DCFCE7; color: #166534 !important; border-radius: 8px; padding: 8px; text-align: center; font-size: .9rem; font-weight: 700; border: 1px solid #86EFAC; margin-bottom: 1rem; }}
.badge-viewer {{ background: #FEF3C7; color: #92400E !important; border-radius: 8px; padding: 8px; text-align: center; font-size: .9rem; font-weight: 700; border: 1px solid #FCD34D; margin-bottom: 1rem; }}
.sidebar-label {{ font-size: 0.85rem; font-weight: 600; color: #94A3B8; margin-bottom: 0.5rem; display: block; }}
.alert-box {{ padding: 12px; border-radius: 10px; margin-bottom: 1rem; font-size: 0.9rem; font-weight: 600; border: 1px solid transparent; }}
.alert-error {{ background-color: #FEF2F2; color: #991B1B; border-color: #FEE2E2; }}
.alert-warning {{ background-color: #FFFBEB; color: #92400E; border-color: #FEF3C7; }}
</style>

""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 💾 Session State
# ──────────────────────────────────────────────
_DEFAULTS = {
    "logged_in":           False,
    "role":                None,
    "assigned_region":     None,
    "login_attempts":      0,
    "locked":              False,
    "preloaded_data":      None,
    "preloaded_passwords": {},
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state: st.session_state[_k] = _v

# ──────────────────────────────────────────────
# ☁️ 雲端服務
# ──────────────────────────────────────────────
@st.cache_resource
def init_supabase() -> Client:
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])

supabase = init_supabase()

@st.cache_data(show_spinner="📥 正在同步數據...")
def download_shared_file(fname: str) -> bytes:
    return supabase.storage.from_(CONFIG["BUCKET_NAME"]).download(fname)

# ──────────────────────────────────────────────
# ⚙️ 工具函數
# ──────────────────────────────────────────────
def safe_div(n, d): return n/d if d and not pd.isna(d) else 0.0

def convert_minguo_date(val):
    try:
        s = str(int(val)).strip()
        y, m = (int(s[:3])+1911, int(s[3:])) if len(s)==5 else (int(s[:2])+1911, int(s[2:]))
        return pd.to_datetime(f"{y}-{m:02d}-01")
    except: return pd.NaT

def apply_chart_style(fig, title="", is_pct=True):
    kw = dict(yaxis_tickformat=".1%") if is_pct else {}
    fig.update_layout(
        title=title, 
        plot_bgcolor=CONFIG["THEME_BG"], # 讓圖表內部背景與網頁一致
        paper_bgcolor=CONFIG["THEME_BG"], # 讓圖表外框背景與網頁一致
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5),
        margin=dict(l=10, r=20, t=40, b=10), **kw
    )
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="rgba(0,0,0,0.05)")
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="rgba(0,0,0,0.05)")
    return fig

# ──────────────────────────────────────────────
# 🧠 決策分析引擎
# ──────────────────────────────────────────────
def classify(eOvd, sOvd, eLoan, shrG, memG):
    T = CONFIG["THRESHOLDS"]
    if eOvd > sOvd and eOvd > T["high_risk_ovd"]: return "🚨 高風險列管"
    if eLoan > T["liquidity_loan"] and shrG < 0: return "⚠️ 流動性緊繃"
    if eLoan < T["idle_loan"] and eOvd < T["ovd_safe_line"]: return "💤 資金閒置"
    if memG > 0 and shrG > 0 and T["stable_loan_min"] < eLoan < T["stable_loan_max"] and eOvd < T["ovd_safe_line"]: return "✅ 穩健模範"
    return "📊 一般狀態"

@st.cache_data(show_spinner="🚀 正在執行智慧分析...")
def process_excel_final(file_bytes: bytes):
    try:
        with pd.ExcelFile(io.BytesIO(file_bytes)) as xls:
            S = CONFIG["SHEETS"]
            if not all(s in xls.sheet_names for s in S.values()):
                raise ValueError("Excel 缺少必要工作表，請檢查分頁名稱。")
            df_m_raw = pd.read_excel(xls, sheet_name=S["MAIN"],   dtype={"社號": str, "年月": str})
            df_l_raw = pd.read_excel(xls, sheet_name=S["LOAN"],   dtype={"社號": str, "年月": str})
            df_r_raw = pd.read_excel(xls, sheet_name=S["REGION"], dtype={"社名": str, "區域": str, "密碼": str})
    except Exception as e: raise ValueError(f"解析失敗: {e}")

    region_map = dict(zip(df_r_raw["社名"], df_r_raw["區域"]))
    region_pw_map = {str(r).strip(): str(p).strip().replace(".0", "") 
                     for r, p in zip(df_r_raw["區域"], df_r_raw["密碼"]) if pd.notna(r) and pd.notna(p)}

    df_m_raw["年月"] = df_m_raw["年月"].apply(convert_minguo_date)
    df_l_raw["年月"] = df_l_raw["年月"].apply(convert_minguo_date)
    for col in ["社員數", "股金", "貸放比"]: df_m_raw[col] = pd.to_numeric(df_m_raw[col], errors="coerce").fillna(0)
    df_m_raw["儲蓄率"] = pd.to_numeric(df_m_raw["儲蓄率"], errors="coerce").fillna(0) / 100
    df_l_raw["逾放比"] = pd.to_numeric(df_l_raw["逾放比"], errors="coerce").fillna(0)
    df_l_raw["收支比"] = pd.to_numeric(df_l_raw["收支比"], errors="coerce").fillna(0) / 100
    if "提撥率" in df_l_raw.columns:
        df_l_raw["提撥率"] = pd.to_numeric(df_l_raw["提撥率"], errors="coerce").fillna(0) / 100
    else:
        df_l_raw["提撥率"] = 0.0

    df_m = df_m_raw.dropna(subset=["年月"]).sort_values(["社號", "年月"])
    df_l = df_l_raw.dropna(subset=["年月"]).sort_values(["社號", "年月"])
    max_d, old_d = df_m["年月"].max(), df_m["年月"].max() - pd.DateOffset(months=12)

    rows = []
    for s_no in df_m["社號"].unique():
        ms, ls = df_m[df_m["社號"] == s_no], df_l[df_l["社號"] == s_no]
        if ms.empty: continue
        name = ms["社名"].iloc[0]
        
        def latest(df, col, d):
            val = df[df["年月"]==d][col].values
            return float(val[0]) if len(val) else float(df.iloc[-1][col])
        
        eM = latest(ms, "社員數", max_d)
        sM = float(ms[ms["年月"]<=old_d]["社員數"].iloc[-1]) if not ms[ms["年月"]<=old_d].empty else float(ms.iloc[0]["社員數"])
        eS = latest(ms, "股金", max_d)
        sS = float(ms[ms["年月"]<=old_d]["股金"].iloc[-1])   if not ms[ms["年月"]<=old_d].empty else float(ms.iloc[0]["股金"])
        
        eOvd = float(ls.iloc[-1]["逾放比"]) if not ls.empty else 0.0
        sOvd = float(ls.iloc[0]["逾放比"]) if not ls.empty else 0.0
        eLoan = float(ms.iloc[-1]["貸放比"])
        memG, shrG = safe_div(eM-sM, sM), safe_div(eS-sS, sS)

        rows.append({
            "社號": s_no, "社名": name, "區域": region_map.get(name, "未分類"),
            "診斷狀態": classify(eOvd, sOvd, eLoan, shrG, memG),
            "現有社員": eM, "社員成長數(12M)": eM - sM, "社員成長率(12M)": memG, "現有股金": eS, "股金成長率(12M)": shrG,
            "貸放比": eLoan, "儲蓄率": float(ms.iloc[-1]["儲蓄率"]),
            "逾放比(初)": sOvd, "逾放比(末)": eOvd, "收支比": float(ls.iloc[-1]["收支比"]) if not ls.empty else 0.0,
            "提撥率": float(ls.iloc[-1]["提撥率"]) if not ls.empty else 0.0,
            "_sM": sM, "_sS": sS
        })
    return pd.DataFrame(rows).fillna(0), df_m, df_l, region_pw_map

# ──────────────────────────────────────────────
# 🔒 存取控管
# ──────────────────────────────────────────────
shared_file = st.query_params.get("file")
if shared_file and st.session_state["preloaded_data"] is None:
    try:
        raw_bytes = download_shared_file(shared_file)
        data, df_m, df_l, region_pws = process_excel_final(raw_bytes)
        st.session_state.update(preloaded_passwords=region_pws, preloaded_data=(data, df_m, df_l, raw_bytes))
    except Exception as e: st.session_state["preload_err"] = str(e)

def handle_login():
    entered = st.session_state.get("pwd_input", "").strip()
    admin_pw = str(st.secrets.get("admin_password", "666"))
    pws = st.session_state.get("preloaded_passwords", {})
    if entered == admin_pw:
        st.session_state.update(logged_in=True, role="admin", assigned_region=None)
    else:
        for r, p in pws.items():
            if entered == p:
                st.session_state.update(logged_in=True, role="viewer", assigned_region=r)
                return
        st.session_state["login_attempts"] += 1
        if st.session_state["login_attempts"] >= CONFIG["MAX_ATTEMPTS"]: st.session_state["locked"] = True

if not st.session_state["logged_in"]:
    _, col, _ = st.columns([0.8, 2.4, 0.8])
    with col:
        with st.container(border=True):
            st.markdown("<h2 class='responsive-h2' style='text-align:center;'>🏦 儲互社分析系統</h2>", unsafe_allow_html=True)
            if st.session_state.get("preload_err"): 
                st.markdown(f'<div class="alert-box alert-error">⚠️ 無法讀取雲端資料，請確認連結。</div>', unsafe_allow_html=True)
            if st.session_state["locked"]: 
                st.markdown(f'<div class="alert-box alert-error">🔒 嘗試次數過多，請稍後再試。</div>', unsafe_allow_html=True)
            else:
                if st.session_state["login_attempts"] > 0: 
                    st.markdown(f'<div class="alert-box alert-warning">❌ 密碼錯誤 ({st.session_state["login_attempts"]}/{CONFIG["MAX_ATTEMPTS"]})</div>', unsafe_allow_html=True)
                st.text_input("密碼", type="password", key="pwd_input", label_visibility="collapsed", placeholder="請輸入密碼")
                st.button("🔓 登入系統", use_container_width=True, on_click=handle_login)
    st.stop()

# ──────────────────────────────────────────────
# 📊 資料載入與過濾
# ──────────────────────────────────────────────
IS_ADMIN = (st.session_state["role"] == "admin")
data_loaded = False

if shared_file and st.session_state["preloaded_data"]:
    data, df_m, df_l, raw_bytes = st.session_state["preloaded_data"]
    region = st.session_state["assigned_region"]
    if region:
        data = data[data["區域"] == region].copy()
        target_snos = data["社號"].unique()
        df_m = df_m[df_m["社號"].isin(target_snos)].copy()
        df_l = df_l[df_l["社號"].isin(target_snos)].copy()
    data_loaded = True
elif IS_ADMIN:
    with st.sidebar:
        st.markdown('<span class="sidebar-label">📂 資料匯入</span>', unsafe_allow_html=True)
        uploaded = st.file_uploader("選擇 Excel 檔案", type=["xlsx"], label_visibility="collapsed")
        if uploaded:
            try:
                raw_bytes = uploaded.getvalue()
                data, df_m, df_l, _ = process_excel_final(raw_bytes)
                data_loaded = True
                st.success("✅ 檔案解析成功")
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown('<span class="sidebar-label">🔗 分享功能</span>', unsafe_allow_html=True)
                if st.button("🚀 生成分享連結", use_container_width=True):
                    fname = f"report_{uuid.uuid4().hex[:10]}.xlsx"
                    supabase.storage.from_(CONFIG["BUCKET_NAME"]).upload(fname, raw_bytes, file_options={"content-type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
                    st.session_state["latest_share_url"] = f"{CONFIG['APP_BASE_URL']}/?file={fname}"
                if "latest_share_url" in st.session_state:
                    st.code(st.session_state["latest_share_url"], language="text")
            except Exception as e: st.error(f"❌ 解析失敗: {e}")

if not data_loaded:
    st.info("👋 歡迎使用分析系統！請由側邊欄上傳 Excel 檔案或點擊分享連結。")
    st.stop()

# ──────────────────────────────────────────────
# 📈 視覺化儀表板
# ──────────────────────────────────────────────
st.markdown(f"<h1 class='responsive-h1'>📊 {st.session_state['assigned_region'] or '全台'} 儲互社分析系統</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<span class="sidebar-label">👤 帳號權限</span>', unsafe_allow_html=True)
    badge_cls = "badge-admin" if IS_ADMIN else "badge-viewer"
    badge_txt = "🔑 管理員模式" if IS_ADMIN else f"👁️ 訪客：{st.session_state['assigned_region']}"
    st.markdown(f'<div class="{badge_cls}">{badge_txt}</div>', unsafe_allow_html=True)
    if st.button("🚪 登出系統", use_container_width=True):
        for k, v in _DEFAULTS.items(): st.session_state[k] = v
        st.rerun()

tab_ov, tab_mx, tab_hc, tab_rp, tab_tr = st.tabs(["📊 經營總覽", "🎯 風險矩陣", "🏥 個社健檢", "📋 報表匯出", "📈 趨勢追蹤"])

with tab_ov:
    c1, c2, c3, c4 = st.columns(4)
    total_mem, total_shr = data["現有社員"].sum(), data["現有股金"].sum()
    prev_mem, prev_shr = data["_sM"].sum(), data["_sS"].sum()
    c1.metric("社員總數", f"{int(total_mem):,}", f"{safe_div(total_mem-prev_mem, prev_mem):.2%}")
    c2.metric("股金總額", f"${total_shr/1e8:.2f} 億", f"{safe_div(total_shr-prev_shr, prev_shr):.2%}")
    c3.metric("平均收支比", f"{data['收支比'].mean():.2%}")
    c4.metric("平均逾放比", f"{data['逾放比(末)'].mean():.2%}")
    st.markdown("### 狀態雷達監控")
    def render_card(title, key, cls):
        names = data[data["診斷狀態"].str.contains(key)]["社名"].tolist()
        st.markdown(f"<div class='stat-card'><div class='card-header {cls}'>{title}</div><div style='padding:10px;'>{' '.join([f'<span class=\"name-tag\">{n}</span>' for n in names]) if names else '無標的'}</div></div>", unsafe_allow_html=True)
    sc1, sc2, sc3, sc4 = st.columns(4)
    with sc1: render_card("🚨 高風險", "高風險", "hdr-red")
    with sc2: render_card("⚠️ 緊繃", "流動性", "hdr-orange")
    with sc3: render_card("💤 閒置", "資金閒置", "hdr-blue")
    with sc4: render_card("✅ 穩健", "穩健", "hdr-green")

with tab_mx:
    T = CONFIG["THRESHOLDS"]
    fig = px.scatter(data, x="貸放比", y="逾放比(末)", color="診斷狀態", size="現有社員", hover_name="社名", height=600, color_discrete_map={
        "🚨 高風險列管": "#EF4444", "⚠️ 流動性緊繃": "#F59E0B", "💤 資金閒置": "#3B82F6", "✅ 穩健模範": "#10B981", "📊 一般狀態": "#94A3B8"
    })
    fig.add_hline(y=T["high_risk_ovd"], line_dash="dot", line_color="red")
    fig.add_vline(x=T["liquidity_loan"], line_dash="dot", line_color="orange")
    apply_chart_style(fig)
    st.plotly_chart(fig, use_container_width=True)

with tab_hc:
    target = st.selectbox("請選擇儲互社", data["社名"].unique())
    if target:
        row = data[data["社名"]==target].iloc[0]
        st.markdown(f"#### 【{target}】 狀態：`{row['診斷狀態']}`")
        KEYS = ["貸放比", "儲蓄率", "逾放比(末)", "收支比", "社員成長率(12M)", "股金成長率(12M)"]
        fig_bar = go.Figure([go.Bar(name=target, x=KEYS, y=[row[k] for k in KEYS], marker_color="#3B82F6"), go.Bar(name="平均", x=KEYS, y=[data[k].mean() for k in KEYS], marker_color="#CBD5E1")])
        apply_chart_style(fig_bar, title="指標對比")
        st.plotly_chart(fig_bar, use_container_width=True)
        cols = st.columns(4)
        for i, (k, v) in enumerate([("現有社員", f"{int(row['現有社員']):,}人"), ("現有股金", f"${row['現有股金']:,.0f}"), ("逾放比", f"{row['逾放比(末)']:.2%}"), ("收支比", f"{row['收支比']:.2%}")]): cols[i].metric(k, v)

with tab_rp:
    fmt = {"現有社員": "{:,}", "社員成長數(12M)": "{:+,.0f}", "現有股金": "${:,.0f}", "社員成長率(12M)": "{:.2%}", "股金成長率(12M)": "{:.2%}", "貸放比": "{:.1%}", "逾放比(初)": "{:.2%}", "逾放比(末)": "{:.2%}", "收支比": "{:.2%}", "提撥率": "{:.2%}"}
    def highlight(row): return ['background-color: #FEF2F2; color: #991B1B; font-weight: bold' if "高風險" in str(row["診斷狀態"]) else '' for _ in row]
    df_export = data.drop(columns=["_sM", "_sS"])
    cols_order = ["社號", "社名", "區域", "診斷狀態", "現有社員", "社員成長數(12M)", "社員成長率(12M)", "現有股金", "股金成長率(12M)", "貸放比", "儲蓄率", "逾放比(初)", "逾放比(末)", "收支比", "提撥率"]
    st.dataframe(df_export[cols_order].style.apply(highlight, axis=1).format(fmt), use_container_width=True, height=600)
    st.download_button("📥 匯出 CSV", df_export[cols_order].to_csv(index=False).encode("utf-8-sig"), "report.csv", "text/csv")

with tab_tr:
    df_all = pd.merge(df_m, df_l[["年月", "社號", "逾放比", "收支比", "提撥率"]], on=["年月", "社號"], how="left")
    sel = st.multiselect("加入比較社別", data["社名"].unique(), [data["社名"].iloc[0]])
    if sel:
        plot_df = df_all[df_all["社名"].isin(sel)]
        def trend(col, title):
            fig = px.line(plot_df, x="年月", y=col, color="社名", markers=True)
            apply_chart_style(fig, title)
            st.plotly_chart(fig, use_container_width=True)
        r1, r2 = st.columns(2)
        with r1: trend("社員數", "👥 社員數趨勢")
        with r2: trend("貸放比", "💰 貸放比趨勢")
        r3, r4 = st.columns(2)
        with r3: trend("儲蓄率", "🏦 儲蓄率趨勢")
        with r4: trend("逾放比", "⚠️ 逾放比趨勢")
        r5, r6 = st.columns(2)
        with r5: trend("收支比", "📈 收支比趨勢")
        with r6: trend("提撥率", "🛡️ 提撥率趨勢")
