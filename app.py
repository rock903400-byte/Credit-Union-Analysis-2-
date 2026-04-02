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
    """安全除法：避免除以 0 或 NaN 造成的運算錯誤"""
    if pd.isna(denominator) or denominator == 0:
        return 0
    result = numerator / denominator
    return result if not pd.isna(result) else 0

# ==========================================
# 1. 頁面基礎設定
# ==========================================
st.set_page_config(page_title="儲互社決策分析中心", layout="wide", page_icon="🏦")

# ==========================================
# 🛑 系統登入密碼鎖 (視覺校正完整版)
# ==========================================
def check_password():
    if st.session_state["password_input"] == str(st.secrets["system_password"]):
        st.session_state["logged_in"] = True
    else:
        st.session_state["logged_in"] = False
        st.error("❌ 密碼錯誤，請重新輸入！")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    # 核心修正：加大 max-width 並確保不斷行，所有元素對齊
    st.markdown("""
        <style>
        /* 強制亮色背景與深色文字 */
        [data-testid="stAppViewContainer"] {
            background-color: #F8FAFC !important;
        }
        .login-wrapper {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 60px 20px;
            max-width: 600px;
            margin: 0 auto;
            text-align: center;
        }
        .login-title {
            color: #1E293B;
            font-size: 2.5rem;
            font-weight: 800;
            margin: 10px 0;
            white-space: nowrap; /* 確保標題不斷行 */
        }
        .login-subtitle {
            color: #64748B;
            font-size: 1.1rem;
            margin-bottom: 30px;
        }
        /* 讓 Streamlit 的輸入框與按鈕符合我們的容器寬度 */
        .stTextInput, .stButton {
            width: 100% !important;
            max-width: 400px !important;
        }
        </style>
        <div class="login-wrapper">
            <div style="font-size: 5rem; margin-bottom: 10px;">🏦</div>
            <div class="login-title">儲互社雲端決策中心</div>
            <div class="login-subtitle">請輸入系統存取密碼以繼續</div>
        </div>
    """, unsafe_allow_html=True)
    
    # 透過排版讓元件出現在上述 wrapper 下方並保持置中
    _, col_mid, _ = st.columns([1, 2, 1])
    with col_mid:
        st.text_input("密碼", type="password", key="password_input", label_visibility="collapsed", placeholder="請輸入密碼")
        st.button("🔓 登入系統", on_click=check_password, use_container_width=True)
    
    st.stop()

# ==========================================
# 🟢 核心程式碼 (登入後執行)
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
    
    /* 調整卡片樣式 */
    .stat-card {
        background: white; border-radius: 12px; border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 1rem;
        min-height: 180px; display: flex; flex-direction: column; overflow: hidden;
    }
    .card-header {
        padding: 12px; color: #FFFFFF !important; font-weight: 700; font-size: 1rem;
        text-align: center; background: #1E293B;
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

@st.cache_data(show_spinner="📊 正在解析決策數據...")
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
        if m_sub.empty: continue
        
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

# --- 側邊欄邏輯 ---
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

uploaded_file = st.sidebar.file_uploader("匯入資料 (Excel)", type=["xlsx"])
if uploaded_file:
    data, df_m, df_l = process_excel_only(uploaded_file)
    data_loaded = True
    if st.sidebar.button("🚀 生成分享連結"):
        fname = f"report_{uuid.uuid4().hex[:8]}.xlsx"
        supabase.storage.from_(BUCKET_NAME).upload(fname, uploaded_file.getvalue())
        st.sidebar.code(f"https://8asdxeziyl2ozfrmkpzof3.streamlit.app/?file={fname}")

# --- 主介面渲染 ---
if data_loaded:
    t1, t2, t3, t4 = st.tabs(["📊 經營總覽", "🎯 風險矩陣", "🏥 個社健檢", "📋 趨勢追蹤"])

    with t1:
        st.markdown("### 🏆 全區指標")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("社員總數", f"{int(data['現有社員'].sum()):,}")
        m2.metric("股金總額", f"${data['現有股金'].sum()/1e8:.2f} 億")
        m3.metric("平均收支比", f"{data['收支比'].mean():.2%}")
        m4.metric("平均逾放比", f"{data['逾放比(末)'].mean():.2%}")

        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        def draw_card(title, s_list, cls):
            tags = "".join([f'<span class="name-tag">{n}</span>' for n in s_list]) if s_list else "無標的"
            st.markdown(f'<div class="stat-card"><div class="card-header {cls}">{title}</div><div class="card-body">{tags}</div></div>', unsafe_allow_html=True)
        
        draw_card("🚨 高風險列管", data[data['診斷狀態']=='🚨 高風險列管']['社名'].tolist(), "header-red")
        with c2: draw_card("⚠️ 流動性緊繃", data[data['診斷狀態']=='⚠️ 流動性緊繃']['社名'].tolist(), "header-orange")
        with c3: draw_card("💤 資金閒置", data[data['診斷狀態']=='💤 資金閒置']['社名'].tolist(), "header-blue")
        with c4: draw_card("✅ 穩健模範", data[data['診斷狀態']=='✅ 穩健模範']['社名'].tolist(), "header-green")

    with t2:
        st.markdown("### 🎯 風險分佈矩陣")
        fig = px.scatter(data, x='貸放比', y='逾放比(末)', color='診斷狀態', size='現有社員', hover_name='社名',
                         color_discrete_map={'🚨 高風險列管':'#EF4444', '⚠️ 流動性緊繃':'#F59E0B', '💤 資金閒置':'#3B82F6', '✅ 穩健模範':'#10B981', '📊 一般狀態':'#94A3B8'},
                         size_max=30, height=600)
        fig.update_layout(xaxis_tickformat='.1%', yaxis_tickformat='.1%', plot_bgcolor="white", margin=dict(l=10, r=10, t=30, b=10), legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig, use_container_width=True)

    with t3:
        target = st.selectbox("選擇儲互社：", data['社名'].unique())
        row = data[data['社名']==target].iloc[0]
        st.subheader(f"診斷結果：{row['診斷狀態']}")
        st.json(row.to_dict())

    with t4:
        st.write("歷史趨勢分析 (開發中)")
else:
    st.info("請於左側上傳檔案以開始分析。")
