"""
儲互社雲端決策中心 — 最終完整版
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
# 全域配置與閾值 (可從 st.secrets 動態調整)
# ──────────────────────────────────────────────
CONFIG = {
    "BUCKET_NAME": "excel-reports",
    "APP_BASE_URL": "https://8asdxeziyl2ozfrmkpzof3.streamlit.app",
    "MAX_LOGIN_ATTEMPTS": 5,
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
# 頁面設定（必須是第一個 st 指令）
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="儲互社決策分析中心",
    layout="wide",
    page_icon="🏦",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────
# 全域 CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600;700&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Noto Sans TC', sans-serif !important;
    background-color: #F0F4F8 !important;
    color: #1A202C !important;
}
h1,h2,h3,h4,h5,h6,p,span,label,
div[data-testid="stMetricValue"],
.stTabs [data-baseweb="tab"] div { color: #1A202C !important; }

/* ── 側邊欄 ── */
[data-testid="stSidebar"] { background-color: #1E293B !important; }
[data-testid="stSidebar"] * { color: #E2E8F0 !important; }
[data-testid="stSidebar"] .stButton > button {
    background: #334155; color: #E2E8F0 !important;
    border: 1px solid #475569; border-radius: 8px;
}
[data-testid="stSidebar"] .stButton > button:hover { background: #475569; }

/* ── 登入卡片：用 st.container(border=True) 的原生容器來做 ── */
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 20px !important;
    border: 1px solid #E2E8F0 !important;
    box-shadow: 0 20px 60px rgba(0,0,0,.09), 0 4px 16px rgba(0,0,0,.05) !important;
    background: white !important;
    padding: .5rem 1.5rem 1.5rem !important;
}

/* ── 狀態指標卡 ── */
.stat-card {
    background: white; border-radius: 14px; border: 1px solid #E2E8F0;
    box-shadow: 0 2px 8px rgba(0,0,0,.06); margin-bottom: 1rem;
    min-height: 180px; display: flex; flex-direction: column; overflow: hidden;
}
.card-header {
    padding: 10px 14px; color: #FFF !important; font-weight: 700; font-size: .95rem;
    display: flex; align-items: center; justify-content: center; gap: 6px;
}
.hdr-red    { background: linear-gradient(135deg, #EF4444, #991B1B); }
.hdr-orange { background: linear-gradient(135deg, #F59E0B, #92400E); }
.hdr-blue   { background: linear-gradient(135deg, #3B82F6, #1E40AF); }
.hdr-green  { background: linear-gradient(135deg, #10B981, #065F46); }
.card-body  { padding: 10px 12px; overflow-y: auto; flex-grow: 1; }
.name-tag {
    display: inline-block; background: #F1F5F9; color: #1A202C !important;
    padding: 3px 10px; border-radius: 8px; margin: 3px;
    font-size: .82rem; border: 1px solid #CBD5E1; font-weight: 600;
}
.no-target { color: #94A3B8 !important; text-align: center; margin-top: 20px; font-size: .85rem; }

/* ── 角色徽章 ── */
.badge-admin {
    background: #DCFCE7; color: #166534 !important; border: 1px solid #86EFAC;
    border-radius: 8px; padding: 6px 10px; font-size: .82rem; text-align: center;
}
.badge-viewer {
    background: #FEF3C7; color: #92400E !important; border: 1px solid #FCD34D;
    border-radius: 8px; padding: 6px 10px; font-size: .82rem; text-align: center;
}

#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
.stTabs [data-baseweb="tab"] { font-size: .95rem; font-weight: 600; padding: 10px 14px; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Session State 統一初始化
# ──────────────────────────────────────────────
_DEFAULTS = {
    "logged_in":           False,
    "role":                None,   # "admin" | "viewer"
    "assigned_region":     None,   # 特定的區域名稱，如 "北區"
    "login_attempts":      0,
    "locked":              False,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ──────────────────────────────────────────────
# Supabase & 雲端快取函數
# ──────────────────────────────────────────────
@st.cache_resource
def init_supabase() -> Client:
    return create_client(
        st.secrets["supabase"]["url"],
        st.secrets["supabase"]["key"],
    )

supabase = init_supabase()

# ──────────────────────────────────────────────
# 工具函數 & 診斷分類
# ──────────────────────────────────────────────
def safe_div(num, den) -> float:
    if pd.isna(den) or den == 0:
        return 0.0
    r = num / den
    return 0.0 if pd.isna(r) else float(r)

def convert_minguo_date(val):
    try:
        s = str(int(val)).strip()
        if   len(s) == 5: y, m = int(s[:3]) + 1911, int(s[3:])
        elif len(s) == 4: y, m = int(s[:2]) + 1911, int(s[2:])
        else: return pd.NaT
        return pd.to_datetime(f"{y}-{m:02d}-01")
    except Exception:
        return pd.NaT

def name_tags(names: list) -> str:
    if not names:
        return '<div class="no-target">無標的</div>'
    return "".join(f'<span class="name-tag">{n}</span>' for n in names)

def stat_card(title: str, names: list, hdr_cls: str):
    st.markdown(
        f'<div class="stat-card">'
        f'<div class="card-header {hdr_cls}">{title}</div>'
        f'<div class="card-body">{name_tags(names)}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

def apply_chart_style(fig, title="", is_pct=True):
    kw = dict(yaxis_tickformat=".1%") if is_pct else {}
    fig.update_layout(
        title=title,
        plot_bgcolor="white",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5),
        margin=dict(l=10, r=20, t=40, b=10),
        **kw,
    )
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="#F1F5F9")
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="#F1F5F9")
    return fig

STATUS = {
    "high_risk": "🚨 高風險列管",
    "liquidity": "⚠️ 流動性緊繃",
    "idle":      "💤 資金閒置",
    "stable":    "✅ 穩健模範",
    "normal":    "📊 一般狀態",
}
COLOR_MAP = {
    STATUS["high_risk"]: "#EF4444",
    STATUS["liquidity"]: "#F59E0B",
    STATUS["idle"]:      "#3B82F6",
    STATUS["stable"]:    "#10B981",
    STATUS["normal"]:    "#94A3B8",
}

def classify(eOvd, sOvd, eLoan, shrG, memG) -> str:
    T = CONFIG["THRESHOLDS"]
    if eOvd > sOvd and eOvd > T["high_risk_ovd"]:           return STATUS["high_risk"]
    if eLoan > T["liquidity_loan"] and shrG < 0:            return STATUS["liquidity"]
    if eLoan < T["idle_loan"] and eOvd < T["ovd_safe_line"]: return STATUS["idle"]
    if memG > 0 and shrG > 0 and T["stable_loan_min"] < eLoan < T["stable_loan_max"] \
       and eOvd < T["ovd_safe_line"]:                        return STATUS["stable"]
    return STATUS["normal"]


# ──────────────────────────────────────────────
# 資料處理引擎
# ──────────────────────────────────────────────
@st.cache_data(show_spinner="🚀 正在啟動決策引擎與數據解析...")
def process_excel(file_bytes: bytes):
    try:
        with pd.ExcelFile(io.BytesIO(file_bytes)) as xls:
            S = CONFIG["SHEETS"]
            missing_sheets = {S["MAIN"], S["LOAN"], S["REGION"]} - set(xls.sheet_names)
            if missing_sheets:
                raise ValueError(f"Excel 缺少必要工作表：{', '.join(missing_sheets)}")
            
            df_m_raw = pd.read_excel(xls, sheet_name=S["MAIN"],   dtype={"社號": str, "年月": str})
            df_l_raw = pd.read_excel(xls, sheet_name=S["LOAN"],   dtype={"社號": str, "年月": str})
            df_r_raw = pd.read_excel(xls, sheet_name=S["REGION"], dtype={"社名": str, "區域": str, "密碼": str})

            required_cols = {
                S["MAIN"]:   ["社號", "社名", "年月", "社員數", "股金", "貸放比", "儲蓄率"],
                S["LOAN"]:   ["社號", "年月", "逾放比", "提撥率", "收支比"],
                S["REGION"]: ["社名", "區域", "密碼"]
            }
            for sheet, cols in required_cols.items():
                target_df = df_m_raw if sheet == S["MAIN"] else (df_l_raw if sheet == S["LOAN"] else df_r_raw)
                missing_cols = set(cols) - set(target_df.columns)
                if missing_cols:
                    raise ValueError(f"工作表 '{sheet}' 缺少必要欄位：{', '.join(missing_cols)}")

    except ValueError as ve:
        raise ve
    except Exception as e:
        raise ValueError(f"檔案解析失敗：{e}") from e

    # 區域映射與密碼字典 (從 Excel 動態取得)
    region_map = dict(zip(df_r_raw["社名"], df_r_raw["區域"]))

    region_pw_map = {}
    for _, row in df_r_raw.dropna(subset=["區域", "密碼"]).iterrows():
        r = str(row["區域"]).strip()
        p = str(row["密碼"]).strip()
        if p.endswith('.0'):  # 處理 Excel 數字轉字串時可能出現的 .0
            p = p[:-2]
        if r != "nan" and p != "nan" and r and p:
            region_pw_map[r] = p

    df_m_raw["年月"] = df_m_raw["年月"].apply(convert_minguo_date)
    df_l_raw["年月"] = df_l_raw["年月"].apply(convert_minguo_date)

    for col in ["社員數", "股金", "貸放比"]:
        df_m_raw[col] = pd.to_numeric(df_m_raw[col], errors="coerce").fillna(0)
    df_m_raw["儲蓄率"] = pd.to_numeric(df_m_raw["儲蓄率"], errors="coerce").fillna(0) / 100
    df_l_raw["逾放比"] = pd.to_numeric(df_l_raw["逾放比"], errors="coerce").fillna(0)
    df_l_raw["提撥率"] = pd.to_numeric(df_l_raw["提撥率"], errors="coerce").fillna(0) / 100
    df_l_raw["收支比"] = pd.to_numeric(df_l_raw["收支比"], errors="coerce").fillna(0) / 100

    df_m = df_m_raw.dropna(subset=["年月"]).sort_values(["社號", "年月"])
    df_l = df_l_raw.dropna(subset=["年月"]).sort_values(["社號", "年月"])

    max_date     = df_m["年月"].max()
    date_12m_ago = max_date - pd.DateOffset(months=12)

    def latest(g, col):
        v = g.loc[g["年月"] == max_date, col].values
        return float(v[0]) if len(v) else float(g.iloc[-1][col])

    def earliest_before(g, col):
        sub = g.loc[g["年月"] <= date_12m_ago, col]
        return float(sub.iloc[-1]) if len(sub) else float(g.iloc[0][col])

    rows = []
    for s_no in df_m["社號"].unique():
        ms   = df_m[df_m["社號"] == s_no]
        ls   = df_l[df_l["社號"] == s_no]
        name = ms["社名"].iloc[0]

        eM, sM = latest(ms, "社員數"), earliest_before(ms, "社員數")
        eS, sS = latest(ms, "股金"),   earliest_before(ms, "股金")
        eOvd   = latest(ls, "逾放比") if not ls.empty else 0.0
        sOvd   = float(ls.iloc[0]["逾放比"]) if not ls.empty else 0.0
        eLoan  = latest(ms, "貸放比")
        memG   = safe_div(eM - sM, sM)
        shrG   = safe_div(eS - sS, sS)

        rows.append({
            "社號": s_no, "社名": name,
            "區域": region_map.get(name, "未分類"),
            "診斷狀態":     classify(eOvd, sOvd, eLoan, shrG, memG),
            "現有社員":     eM,    "社員成長率(12M)": memG,
            "現有股金":     eS,    "股金成長率(12M)": shrG,
            "貸放比":       eLoan, "儲蓄率":          latest(ms, "儲蓄率"),
            "逾放比(初)":  sOvd,  "逾放比(末)":      eOvd,
            "提撥率":       latest(ls, "提撥率") if not ls.empty else 0.0,
            "收支比":       latest(ls, "收支比") if not ls.empty else 0.0,
            "_sM": sM, "_sS": sS,
        })

    return pd.DataFrame(rows).fillna(0), df_m, df_l, region_pw_map


# ──────────────────────────────────────────────
# 預先載入分享資料 (供訪客登入比對密碼)
# ──────────────────────────────────────────────
shared_file = st.query_params.get("file")

if shared_file and st.session_state.get("preloaded_data") is None:
    try:
        raw_bytes = download_shared_file(shared_file)
        data, df_m, df_l, region_pw_map = process_excel(raw_bytes)
        st.session_state["preloaded_passwords"] = region_pw_map
        st.session_state["preloaded_data"] = (data, df_m, df_l, raw_bytes)
    except Exception as e:
        logger.error("Cloud load failed:\n%s", traceback.format_exc())

# ──────────────────────────────────────────────
# 🔐 登入邏輯
# ──────────────────────────────────────────────
def handle_login():
    if st.session_state["locked"]:
        return
    entered      = st.session_state.get("pwd_input", "").strip()
    admin_pw     = str(st.secrets.get("admin_password", ""))
    
    # 從預先載入的 Excel 字典取得區域密碼
    regional_pws = st.session_state.get("preloaded_passwords", {})

    if entered == admin_pw:
        st.session_state.update(logged_in=True, role="admin", assigned_region=None, login_attempts=0)
    else:
        matched_region = None
        for region, pw in regional_pws.items():
            # 確保比對時忽略可能的結尾 .0 或空白
            target_pw = str(pw).strip()
            if target_pw.endswith('.0'):
                target_pw = target_pw[:-2]
                
            if entered == target_pw:
                matched_region = region
                break
        
        if matched_region:
            st.session_state.update(logged_in=True, role="viewer", assigned_region=matched_region, login_attempts=0)
        else:
            st.session_state["login_attempts"] += 1
            if st.session_state["login_attempts"] >= CONFIG["MAX_LOGIN_ATTEMPTS"]:
                st.session_state["locked"] = True



def render_login():
    """登入頁：三欄置中 + st.container(border=True) 卡片"""
    st.markdown("<div style='margin-top:6vh'></div>", unsafe_allow_html=True)

    _, center, _ = st.columns([1, 1.6, 1])
    with center:
        with st.container(border=True):
            st.markdown("""
                <div style="text-align:center; padding:1.5rem 0 .8rem;">
                    <div style="font-size:3.8rem; line-height:1;">🏦</div>
                    <h2 style="font-size:1.75rem; font-weight:700;
                               color:#1A202C; margin:.4rem 0 .2rem;">
                        儲互社雲端決策中心
                    </h2>
                    <p style="color:#64748B; font-size:.92rem; margin:0;">
                        請輸入系統存取密碼以繼續
                    </p>
                </div>
            """, unsafe_allow_html=True)

            if st.session_state["locked"]:
                st.error(
                    f"🔒 嘗試次數超過 {MAX_LOGIN_ATTEMPTS} 次，"
                    "請重新整理頁面後再試。",
                    icon=None,
                )
            else:
                attempts  = st.session_state["login_attempts"]
                remaining = CONFIG["MAX_LOGIN_ATTEMPTS"] - attempts
                if attempts > 0:
                    st.warning(f"⚠️ 密碼錯誤，剩餘嘗試次數：{remaining} 次")

                st.text_input(
                    "密碼", type="password", key="pwd_input",
                    label_visibility="collapsed",
                    placeholder="請輸入密碼",
                    disabled=st.session_state["locked"],
                )
                st.button(
                    "🔓 登入系統",
                    on_click=handle_login,
                    use_container_width=True,
                    disabled=st.session_state["locked"],
                )

    st.stop()


if not st.session_state["logged_in"]:
    render_login()


# ──────────────────────────────────────────────
# 登入後常數
# ──────────────────────────────────────────────
IS_ADMIN = (st.session_state["role"] == "admin")


# ──────────────────────────────────────────────
# 側邊欄
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ 決策數據中心")
    badge_cls = "badge-admin"  if IS_ADMIN else "badge-viewer"
    badge_txt = "🔑 管理員模式（可上傳）" if IS_ADMIN else "👁️ 訪客模式（僅限瀏覽）"
    st.markdown(f'<div class="{badge_cls}">{badge_txt}</div>', unsafe_allow_html=True)
    st.markdown("---")
    if st.button("🚪 登出系統", use_container_width=True):
        for k, v in _DEFAULTS.items():
            st.session_state[k] = v
        # 清除 URL 參數，避免登出後重新觸發自動載入
        st.query_params.clear()
        st.rerun()
    st.markdown("---")


# ──────────────────────────────────────────────
# 資料載入與過濾
# ──────────────────────────────────────────────
data_loaded = False
data = df_m = df_l = None
raw_bytes = None

shared_file = st.query_params.get("file")

def apply_region_filter(df_main, df_m_raw, df_l_raw):
    """根據權限過濾資料"""
    region = st.session_state.get("assigned_region")
    if region:
        # 過濾主表
        df_main = df_main[df_main["區域"] == region].copy()
        # 過濾原始趨勢表
        target_s_nos = df_main["社號"].unique()
        df_m_raw = df_m_raw[df_m_raw["社號"].isin(target_s_nos)].copy()
        df_l_raw = df_l_raw[df_l_raw["社號"].isin(target_s_nos)].copy()
    return df_main, df_m_raw, df_l_raw

if shared_file:
    with st.sidebar:
        st.info("📁 正在從雲端載入...")
    try:
        raw_bytes = download_shared_file(shared_file)
        data, df_m, df_l = process_excel(raw_bytes)
        data, df_m, df_l = apply_region_filter(data, df_m, df_l)
        data_loaded = True
        with st.sidebar:
            st.success("✅ 雲端資料載入成功！")
            if st.session_state.get("assigned_region"):
                st.warning(f"📍 已套用區域過濾：{st.session_state['assigned_region']}")
    except Exception as e:
        with st.sidebar:
            st.error(f"❌ 連結失效或解析錯誤 (Debug: {e})")
        logger.error("Cloud load failed:\n%s", traceback.format_exc())

elif IS_ADMIN:
    with st.sidebar:
        uploaded = st.file_uploader("📂 匯入 Excel 檔案", type=["xlsx"])

    if uploaded:
        try:
            raw_bytes = uploaded.getvalue()
            # 管理員上傳時，忽略回傳的密碼字典
            data, df_m, df_l, _ = process_excel(raw_bytes)
            data_loaded = True
            with st.sidebar:
                st.success("✅ 檔案解析成功！")
                st.markdown("---")
                st.markdown("### ☁️ 雲端分享功能")
                if st.button("🚀 生成分享連結", use_container_width=True):
                    with st.spinner("上傳中..."):
                        try:
                            fname = f"report_{uuid.uuid4().hex[:10]}.xlsx"
                            supabase.storage.from_(CONFIG["BUCKET_NAME"]).upload(
                                fname, raw_bytes,
                                file_options={
                                    "content-type": (
                                        "application/vnd.openxmlformats-"
                                        "officedocument.spreadsheetml.sheet"
                                    ),
                                    "x-upsert": "true",
                                },
                            )
                            url = f"{CONFIG['APP_BASE_URL']}/?file={fname}"
                            st.success("✅ 上傳成功！")
                            st.code(url, language="text")
                            st.caption("將此連結分享給訪客即可直接瀏覽。")
                        except Exception as e:
                            st.error(f"上傳失敗：{e}")
                            logger.error("Upload failed:\n%s", traceback.format_exc())
        except ValueError as e:
            with st.sidebar:
                st.error(f"❌ {e}")
        except Exception:
            with st.sidebar:
                st.error("❌ 未知錯誤，請確認檔案格式正確。")
            logger.error("Excel parse failed:\n%s", traceback.format_exc())

else:
    with st.sidebar:
        st.info("📎 請使用管理員提供的分享連結載入報表資料。")


# ──────────────────────────────────────────────
# 尚未載入時的歡迎畫面
# ──────────────────────────────────────────────
if not data_loaded:
    st.markdown("""
        <div style="text-align:center; margin-top:12vh; color:#64748B;">
            <div style="font-size:4rem;">🏦</div>
            <h1 style="color:#1A202C; font-size:2rem; margin:.5rem 0;">
                儲互社雲端決策中心
            </h1>
            <p style="font-size:1rem;">
                管理員請於左側上傳 Excel 檔案　／　訪客請使用分享連結直接載入。
            </p>
        </div>
    """, unsafe_allow_html=True)
    st.stop()


# ──────────────────────────────────────────────
# 主畫面 Tabs
# ──────────────────────────────────────────────
tab_ov, tab_mx, tab_hc, tab_rp, tab_tr = st.tabs([
    "📊 經營總覽", "🎯 全域風險矩陣", "🏥 個社健檢", "📋 報表匯出", "📈 趨勢追蹤",
])


# ════════════════════════════════════
# Tab 1 ： 經營總覽
# ════════════════════════════════════
with tab_ov:
    st.markdown("### 🏆 區會總體指標")
    total_mem = data["現有社員"].sum()
    total_shr = data["現有股金"].sum()
    prev_mem  = data["_sM"].sum()
    prev_shr  = data["_sS"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("全體社員總數",   f"{int(total_mem):,}",
              f"{safe_div(total_mem - prev_mem, prev_mem):.2%}")
    c2.metric("全體股金總額",   f"${total_shr / 1e8:.2f} 億",
              f"{safe_div(total_shr - prev_shr, prev_shr):.2%}")
    c3.metric("全區平均收支比", f"{data['收支比'].mean():.2%}")
    c4.metric("全區平均逾放比", f"{data['逾放比(末)'].mean():.2%}")

    st.markdown("### 🏷️ 狀態雷達監控")

    def names_of(key): return data[data["診斷狀態"] == STATUS[key]]["社名"].tolist()

    sc1, sc2, sc3, sc4 = st.columns(4)
    with sc1: stat_card("🚨 高風險列管", names_of("high_risk"), "hdr-red")
    with sc2: stat_card("⚠️ 流動性緊繃", names_of("liquidity"), "hdr-orange")
    with sc3: stat_card("💤 資金閒置",   names_of("idle"),      "hdr-blue")
    with sc4: stat_card("✅ 穩健模範",   names_of("stable"),    "hdr-green")


# ════════════════════════════════════
# Tab 2 ： 全域風險矩陣
# ════════════════════════════════════
with tab_mx:
    st.markdown("### 🎯 全區風險分佈矩陣")
    st.caption("氣泡大小 = 社員數。十字線為閾值邊界，協助快速定位高風險社別。")

    fig = px.scatter(
        data, x="貸放比", y="逾放比(末)",
        color="診斷狀態", size="現有社員", hover_name="社名",
        color_discrete_map=COLOR_MAP, size_max=28, height=560,
    )
    fig.add_hline(y=0.10, line_dash="dot", line_color="#EF4444", annotation_text="高風險 10%")
    fig.add_vline(x=0.90, line_dash="dot", line_color="#F59E0B", annotation_text="緊繃 90%")
    fig.add_vline(x=0.30, line_dash="dot", line_color="#3B82F6", annotation_text="閒置 30%")
    apply_chart_style(fig)
    fig.update_layout(
        legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5)
    )
    st.plotly_chart(fig, use_container_width=True)


# ════════════════════════════════════
# Tab 3 ： 個社健檢
# ════════════════════════════════════
with tab_hc:
    st.markdown("### 🏥 單一儲互社健檢報告")
    target = st.selectbox("請選擇要診斷的儲互社：", data["社名"].unique())

    if target:
        row  = data[data["社名"] == target].iloc[0]
        gavg = data.mean(numeric_only=True)
        st.markdown(f"#### 【{target}】目前狀態：`{row['診斷狀態']}`")
        st.markdown("---")

        KEYS   = ["貸放比", "儲蓄率", "逾放比(末)", "收支比", "社員成長率(12M)", "股金成長率(12M)"]
        LABELS = ["貸放比", "儲蓄率", "逾放比(末)", "收支比", "社員成長率",       "股金成長率"]

        fig_bar = go.Figure([
            go.Bar(name=target,    x=LABELS, y=[row[k]  for k in KEYS], marker_color="#3B82F6"),
            go.Bar(name="全區平均", x=LABELS, y=[gavg[k] for k in KEYS], marker_color="#CBD5E1"),
        ])
        apply_chart_style(fig_bar, title="關鍵指標 vs 全區基準")
        fig_bar.update_layout(barmode="group", height=440)
        st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown("##### 📋 詳細數值")
        kv = [
            ("現有社員",   f"{int(row['現有社員']):,} 人"),
            ("現有股金",   f"${row['現有股金']:,.0f}"),
            ("貸放比",     f"{row['貸放比']:.1%}"),
            ("儲蓄率",     f"{row['儲蓄率']:.2%}"),
            ("逾放比(末)", f"{row['逾放比(末)']:.2%}"),
            ("收支比",     f"{row['收支比']:.2%}"),
            ("社員成長率", f"{row['社員成長率(12M)']:.2%}"),
            ("股金成長率", f"{row['股金成長率(12M)']:.2%}"),
        ]
        kv_cols = st.columns(4)
        for i, (lbl, val) in enumerate(kv):
            kv_cols[i % 4].metric(lbl, val)


# ════════════════════════════════════
# Tab 4 ： 報表匯出
# ════════════════════════════════════
with tab_rp:
    st.markdown("### 📋 完整診斷數據總表")
    display_df = data.drop(columns=["_sM", "_sS"])
    fmt = {
        "社員成長率(12M)": "{:.2%}", "股金成長率(12M)": "{:.2%}",
        "貸放比":    "{:.1%}", "儲蓄率":    "{:.2%}",
        "逾放比(初)": "{:.2%}", "逾放比(末)": "{:.2%}",
        "提撥率":    "{:.2%}", "收支比":    "{:.2%}",
        "現有社員":  "{:,}",  "現有股金":  "${:,.0f}",
    }

    def row_highlight(row):
        s = ("background-color:#FEF2F2;color:#991B1B;font-weight:bold;"
             if "高風險" in str(row.get("診斷狀態", "")) else "")
        return [s] * len(row)

    st.dataframe(
        display_df.style.apply(row_highlight, axis=1).format(fmt),
        use_container_width=True, height=560,
    )
    st.download_button(
        "📥 匯出完整診斷報告 (CSV)",
        display_df.to_csv(index=False).encode("utf-8-sig"),
        "診斷報告.csv", "text/csv",
        use_container_width=True,
    )


# ════════════════════════════════════
# Tab 5 ： 趨勢追蹤
# ════════════════════════════════════
with tab_tr:
    st.markdown("### 📈 歷史趨勢對比")

    df_all = pd.merge(
        df_m,
        df_l[["年月", "社號", "逾放比", "收支比"]],
        on=["年月", "社號"], how="left",
    )
    avg_df         = df_all.groupby("年月").mean(numeric_only=True).reset_index()
    avg_df["社名"] = "—— 區域基準 ——"
    BASELINE       = {"—— 區域基準 ——": "#1E293B"}

    show_avg = st.checkbox("顯示區域基準線（黑色虛線）", value=True)
    selected = st.multiselect(
        "加入比較的儲互社：",
        options=data["社名"].unique(),
        default=[data["社名"].iloc[0]],
    )

    if not selected:
        st.info("請至少選擇一家儲互社。")
    else:
        base = df_all[df_all["社名"].isin(selected)]
        plot = pd.concat([base, avg_df]) if show_avg else base

        def trend_chart(col, title, is_pct=True):
            fig = px.line(
                plot, x="年月", y=col, color="社名",
                title=title, markers=True, color_discrete_map=BASELINE,
            )
            fig.for_each_trace(
                lambda t: t.update(line=dict(dash="dash", width=2.5))
                if t.name == "—— 區域基準 ——" else None
            )
            apply_chart_style(fig, title=title, is_pct=is_pct)
            st.plotly_chart(fig, use_container_width=True)

        r1, r2 = st.columns(2)
        with r1: trend_chart("社員數", "👥 社員數趨勢", is_pct=False)
        with r2: trend_chart("貸放比", "💰 貸放比趨勢")
        r3, r4 = st.columns(2)
        with r3: trend_chart("儲蓄率", "🏦 儲蓄率趨勢")
        with r4: trend_chart("逾放比", "⚠️ 逾放比趨勢")
        st.divider()
        trend_chart("收支比", "📈 收支比趨勢")
