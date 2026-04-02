import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import uuid
from supabase import create_client, Client

# ==========================================
# 🛠️ 輔助工具函數
# ==========================================
def safe_div(numerator, denominator):
    if pd.isna(denominator) or denominator == 0:
        return 0
    result = numerator / denominator
    return result if not pd.isna(result) else 0

# ==========================================
# 1. 頁面基礎設定
# ==========================================
st.set_page_config(page_title="儲互社決策分析中心", layout="wide", page_icon="🏦")

# ==========================================
# 🛑 系統登入密碼鎖 (視覺絕對對齊版)
# ==========================================
def check_password():
    if st.session_state["password_input"] == str(st.secrets["system_password"]):
        st.session_state["logged_in"] = True
    else:
        st.session_state["logged_in"] = False
        st.error("❌ 密碼錯誤")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    # 🌟 核心修正：將所有元件強制約束在一個固定寬度的中央容器內
    st.markdown("""
        <style>
        /* 1. 全域背景設定 */
        [data-testid="stAppViewContainer"] {
            background-color: #F8FAFC !important;
        }
        
        /* 2. 定義一個中央登入盒子 */
        .auth-box {
            max-width: 420px;
            margin: 0 auto;
            text-align: center;
            padding: 40px 10px;
        }
        
        /* 3. 確保標題不斷行，顏色統一 */
        .auth-title {
            color: #1E293B !important;
            font-size: 2.5rem !important;
            font-weight: 800 !important;
            margin-bottom: 0.5rem !important;
            white-space: nowrap;
        }
        
        /* 4. 強制覆蓋 Streamlit 元件寬度，使其與容器完全對齊 */
        div[data-testid="stForm"], .stTextInput, .stButton button {
            width: 100% !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # 使用 columns 來定位中央區域
    _, center_col, _ = st.columns([1, 2, 1])
    
    with center_col:
        # 標題區塊
        st.markdown("""
            <div style="text-align: center; margin-top: 5vh;">
                <div style="font-size: 5rem; margin-bottom: 10px;">🏦</div>
                <h1 class="auth-title">儲互社雲端決策中心</h1>
                <p style="color: #64748B; font-size: 1.1rem; margin-bottom: 2rem;">請輸入系統存取密碼以繼續</p>
            </div>
        """, unsafe_allow_html=True)
        
        # 密碼輸入與按鈕 (放在同一個容器內確保寬度一致)
        st.text_input("密碼", type="password", key="password_input", label_visibility="collapsed", placeholder="請輸入密碼")
        st.button("🔓 登入系統", on_click=check_password, use_container_width=True)
    
    st.stop()

# ==========================================
# 🟢 核心程式碼 (登入成功後)
# ==========================================

@st.cache_resource
def init_supabase() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase = init_supabase()
BUCKET_NAME = "excel-reports"

# --- 全域樣式優化 ---
st.markdown("""
    <style>
    [data-testid="stAppViewContainer"] { background-color: #F8FAFC !important; color: #1E293B !important; }
    h1, h2, h3, h4, p, span, label, div[data-testid="stMetricValue"] { color: #1E293B !important; }
    
    .stat-card {
        background: white; border-radius: 12px; border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 1rem;
        min-height: 180px; display: flex; flex-direction: column; overflow: hidden;
    }
    .card-header {
        padding: 12px; color: #FFFFFF !important; font-weight: 700; font-size: 1rem;
        text-align: center;
    }
    .header-red { background: linear-gradient(135deg, #EF4444, #991B1B); }
    .header-orange { background: linear-gradient(135deg, #F59E0B, #92400E); }
    .header-blue { background: linear-gradient(135deg, #3B82F6, #1E40AF); }
    .header-green { background: linear-gradient(135deg, #10B981, #065F46); }
    
    .card-body { padding: 15px; background: #FFFFFF; flex-grow: 1; }
    .name-tag {
        display: inline-block; background: #F1F5F9; color: #1E293B !important; padding: 4px 10px;
        border-radius: 8px; margin: 4px; font-size: 0.85rem; border: 1px solid #CBD5E1; font-weight: 600;
    }
    #MainMenu, footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- 資料解析邏輯 ---
def convert_minguo_date(val):
    try:
        s = str(int(val)).strip()
        if len(s) == 5: year, month = int(s[:3]) + 1911, int(s[3:])
        elif len(s) == 4: year, month = int(s[:2]) + 1911, int(s[2:])
        else: return pd.NaT
        return pd.to_datetime(f"{year}-{month}-01")
    except: return pd.NaT

@st.cache_data(show_spinner="📊 正在解析數據...")
def process_excel_only(file):
    df_m_raw = pd.read_excel(file, sheet_name="社務及資金運用情形", dtype={'社號': str, '年月': str})
    df_l_raw = pd.read_excel(file, sheet_name="放款及逾期放款", dtype={'社號': str, '年月': str})
    df_m_raw['年月'] = df_m_raw['年月'].apply(convert_minguo_date)
    df_l_raw['年月'] = df_l_raw['年月'].apply(convert_minguo_date)
    
    for col in ['社員數', '股金', '貸放比']: 
        df_m_raw[col] = pd.to_numeric(df_m_raw[col], errors='coerce').fillna(0)
    
    df_m = df_m_raw.dropna(subset=['年月']).sort_values(by=['社號', '年月'])
    df_l = df_l_raw.dropna(subset=['年月']).sort_values(by=['社號', '年月'])
    max_date = df_m['年月'].max()
    date_12m_ago = max_date - pd.DateOffset(months=12)
    
    rows = []
    for s_no in df_m['社號'].unique():
        m_sub = df_m[df_m['社號'] == s_no]
        l_sub = df_l[df_l['社號'] == s_no]
        
        def get_v(g, c, lat=True):
            if g.empty: return 0
            v = g[g['年月'] == max_date][c].values if lat else g[g['年月'] <= date_12m_ago].tail(1)[c].values
            return v[0] if len(v)>0 else g.iloc[-1 if lat else 0][c]

        eM, sM = get_v(m_sub, '社員數', True), get_v(m_sub, '社員數', False)
        eS, sS = get_v(m_sub, '股金', True), get_v(m_sub, '股金', False)
        eOverdue = get_v(l_sub, '逾放比', True) if not l_sub.empty else 0
        eLoanRatio = get_v(m_sub, '貸放比', True)
        
        memGrowth = safe_div((eM - sM), sM)
        shrGrowth = safe_div((eS - sS), sS)
        
        status = "📊 一般狀態"
        if eOverdue > 0.1: status = "🚨 高風險列管"
        elif eLoanRatio > 0.9: status = "⚠️ 流動性緊繃"
        elif eLoanRatio < 0.3: status = "💤 資金閒置"
        elif memGrowth > 0 and eLoanRatio > 0.5: status = "✅ 穩健模範"

        rows.append({
            '社號': s_no, '社名': m_sub['社名'].iloc[0], '診斷狀態': status,
            '現有社員': eM, '社員成長率(12M)': memGrowth,
            '現有股金': eS, '股金成長率(12M)': shrGrowth,
            '貸放比': eLoanRatio, '逾放比(末)': eOverdue,
            '收支比': get_v(l_sub, '收支比', True) if not l_sub.empty else 0,
            'sM_total': sM, 'sS_total': sS
        })
    return pd.DataFrame(rows), df_m, df_l

# --- 資料載入 ---
st.sidebar.title("🛠️ 數據管理")
if st.sidebar.button("🚪 登出系統"):
    st.session_state["logged_in"] = False
    st.rerun()

query_params = st.query_params
shared_file = query_params.get("file")
data_loaded = False

if shared_file:
    try:
        res = supabase.storage.from_(BUCKET_NAME).download(shared_file)
        data, df_m, df_l = process_excel_only(io.BytesIO(res))
        data_loaded = True
    except:
        st.error("分享連結已失效")

uploaded_file = st.sidebar.file_uploader("匯入 Excel", type=["xlsx"])
if uploaded_file:
    data, df_m, df_l = process_excel_only(uploaded_file)
    data_loaded = True

# --- 介面渲染 ---
if data_loaded:
    t1, t2, t3 = st.tabs(["📊 總覽", "🎯 矩陣", "🏥 健檢"])
    # (後續功能代碼保持不變...)
    with t1:
        st.markdown("### 🏆 區會指標")
        m1, m2 = st.columns(2)
        m1.metric("社員總數", f"{int(data['現有社員'].sum()):,}")
        m2.metric("股金總額", f"${data['現有股金'].sum()/1e8:.2f} 億")
        
        st.divider()
        st.dataframe(data[['社名', '診斷狀態', '現有社員', '貸放比', '逾放比(末)']], use_container_width=True)
else:
    st.info("請上傳資料以進行分析")
