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
# 🛑 系統登入密碼鎖 (對齊優化版)
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
    # 增加針對登入框的專屬 CSS，強制限制電腦版寬度並置中
    st.markdown("""
        <style>
        .login-container {
            max-width: 450px;
            margin: 0 auto;
            padding: 20px;
            text-align: center;
        }
        /* 修正按鈕與輸入框在電腦版太長的問題 */
        [data-testid="column"] {
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        </style>
        <div class="login-container" style="margin-top: 8vh;">
            <h1 style="font-size: clamp(3.5rem, 10vw, 4.5rem); margin-bottom: 0.5rem;">🏦</h1>
            <h1 style="color: #1E293B; font-size: clamp(2rem, 6vw, 2.8rem); font-weight: 800; margin-bottom: 1rem;">儲互社雲端決策中心</h1>
            <p style="color: #64748B; font-size: 1.1rem; margin-bottom: 2rem;">請輸入系統存取密碼以繼續</p>
        </div>
    """, unsafe_allow_html=True)
    
    # 使用 1:2:1 比例夾住，但在 CSS 中我們加了 max-width 保護
    _, col2, _ = st.columns([1, 2, 1])
    with col2:
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

# --- 自定義 CSS (手機與桌機圖表優化) ---
st.markdown("""
    <style>
    [data-testid="stAppViewContainer"] {
        background-color: #F8FAFC !important;
        color: #1E293B !important;
    }
    h1, h2, h3, h4, h5, h6, p, span, label, div[data-testid="stMetricValue"], .stTabs [data-baseweb="tab"] div {
        color: #1E293B !important;
    }
    .stat-card {
        background: white; border-radius: 12px; border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 1rem;
        min-height: 180px; height: auto; display: flex; flex-direction: column; overflow: hidden;
    }
    .card-header {
        padding: 10px; color: #FFFFFF !important; font-weight: 700; font-size: 1rem;
        text-align: center; display: flex; align-items: center; justify-content: center; gap: 8px;
    }
    .header-red { background: linear-gradient(135deg, #EF4444, #991B1B); }
    .header-orange { background: linear-gradient(135deg, #F59E0B, #92400E); }
    .header-blue { background: linear-gradient(135deg, #3B82F6, #1E40AF); }
    .header-green { background: linear-gradient(135deg, #10B981, #065F46); }
    .card-body { padding: 12px; overflow-y: auto; flex-grow: 1; background: #FFFFFF; }
    .name-tag {
        display: inline-block; background: #F1F5F9; color: #1E293B !important; padding: 4px 10px;
        border-radius: 8px; margin: 3px; font-size: 0.85rem; border: 1px solid #CBD5E1; font-weight: 600;
    }
    .stTabs [data-baseweb="tab"] { font-size: 1.1rem; font-weight: 600; padding: 10px 15px; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- 資料處理與轉換功能 ---
def convert_minguo_date(val):
    try:
        s = str(int(val)).strip()
        if len(s) == 5: year, month = int(s[:3]) + 1911, int(s[3:])
        elif len(s) == 4: year, month = int(s[:2]) + 1911, int(s[2:])
        else: return pd.NaT
        return pd.to_datetime(f"{year}-{month}-01")
    except: return pd.NaT

@st.cache_data(show_spinner="🚀 正在分析數據...")
def process_excel_only(file):
    df_m_raw = pd.read_excel(file, sheet_name="社務及資金運用情形", dtype={'社號': str, '年月': str})
    df_l_raw = pd.read_excel(file, sheet_name="放款及逾期放款", dtype={'社號': str, '年月': str})
    df_m_raw['年月'] = df_m_raw['年月'].apply(convert_minguo_date)
    df_l_raw['年月'] = df_l_raw['年月'].apply(convert_minguo_date)
    for col in ['社員數', '股金', '貸放比']: df_m_raw[col] = pd.to_numeric(df_m_raw[col], errors='coerce').fillna(0)
    df_m_raw['儲蓄率'] = pd.to_numeric(df_m_raw['儲蓄率'], errors='coerce').fillna(0) / 100
    df_l_raw['逾放比'] = pd.to_numeric(df_l_raw['逾放比'], errors='coerce').fillna(0)
    df_l_raw['提撥率'] = pd.to_numeric(df_l_raw['提撥率'], errors='coerce').fillna(0) / 100
    df_l_raw['收支比'] = pd.to_numeric(df_l_raw['收支比'], errors='coerce').fillna(0) / 100
    df_m = df_m_raw.dropna(subset=['年月']).sort_values(by=['社號', '年月'])
    df_l = df_l_raw.dropna(subset=['年月']).sort_values(by=['社號', '年月'])
    max_date = df_m['年月'].max()
    date_12m_ago = max_date - pd.DateOffset(months=12)
    societies = df_m['社號'].unique()
    rows = []
    for s_no in societies:
        m_sub = df_m[df_m['社號'] == s_no]
        l_sub = df_l[df_l['社號'] == s_no]
        s_name = m_sub['社名'].iloc[0]
        def get_v(g, c, lat=True):
            if g.empty: return 0
            if lat:
                v = g[g['年月'] == max_date][c].values
                return v[0] if len(v)>0 else g.iloc[-1][c]
            else:
                v = g[g['年月'] <= date_12m_ago].tail(1)[c].values
                return v[0] if len(v)>0 else g.iloc[0][c]
        eM, sM = get_v(m_sub, '社員數', True), get_v(m_sub, '社員數', False)
        eS, sS = get_v(m_sub, '股金', True), get_v(m_sub, '股金', False)
        eOverdue = get_v(l_sub, '逾放比', True)
        sOverdue = l_sub.iloc[0]['逾放比'] if not l_sub.empty else 0
        eLoanRatio = get_v(m_sub, '貸放比', True)
        memGrowth = safe_div((eM - sM), sM)
        shrGrowth = safe_div((eS - sS), sS)
        status = "📊 一般狀態"
        if eOverdue > sOverdue and eOverdue > 0.1: status = "🚨 高風險列管"
        elif eLoanRatio > 0.9 and shrGrowth < 0: status = "⚠️ 流動性緊繃"
        elif eLoanRatio < 0.3 and eOverdue < 0.02: status = "💤 資金閒置"
        elif memGrowth > 0 and shrGrowth > 0 and 0.4 < eLoanRatio < 0.8 and eOverdue < 0.02: status = "✅ 穩健模範"
        rows.append({
            '社號': s_no, '社名': s_name, '診斷狀態': status,
            '現有社員': eM, '社員成長率(12M)': memGrowth,
            '現有股金': eS, '股金成長率(12M)': shrGrowth,
            '貸放比': eLoanRatio, '儲蓄率': get_v(m_sub, '儲蓄率', True),
            '逾放比(初)': sOverdue, '逾放比(末)': eOverdue,
            '提撥率': get_v(l_sub, '提撥率', True), '收支比': get_v(l_sub, '收支比', True),
            'sM_total': sM, 'sS_total': sS
        })
    return pd.DataFrame(rows).fillna(0), df_m, df_l

# --- 側邊欄與載入 ---
st.sidebar.markdown("## ⚙️ 決策數據中心")
if st.sidebar.button("🚪 登出系統"):
    st.session_state["logged_in"] = False
    st.rerun()

st.sidebar.markdown("---")
query_params = st.query_params
shared_file = query_params.get("file")
data_loaded = False

if shared_file:
    try:
        res = supabase.storage.from_(BUCKET_NAME).download(shared_file)
        data, df_m, df_l = process_excel_only(io.BytesIO(res))
        data_loaded = True
        st.sidebar.success("✅ 雲端載入成功")
    except:
        st.sidebar.error("❌ 連結已失效")
else:
    uploaded_file = st.sidebar.file_uploader("匯入 Excel 檔案", type=["xlsx"])
    if uploaded_file:
        try:
            data, df_m, df_l = process_excel_only(uploaded_file)
            data_loaded = True
            st.sidebar.success("✅ 檔案解析成功")
            if st.sidebar.button("🚀 生成分享連結"):
                safe_filename = f"report_{uuid.uuid4().hex[:8]}.xlsx"
                supabase.storage.from_(BUCKET_NAME).upload(safe_filename, uploaded_file.getvalue())
                st.sidebar.code(f"https://8asdxeziyl2ozfrmkpzof3.streamlit.app/?file={safe_filename}")
        except:
            st.sidebar.error("❌ 解析失敗，請檢查格式")

# --- 介面渲染 ---
if data_loaded:
    t1, t2, t3, t4, t5 = st.tabs(["📊 經營總覽", "🎯 風險矩陣", "🏥 個社健檢", "📋 診斷總表", "📈 趨勢追蹤"])

    with t1:
        st.markdown("### 🏆 區會總體指標")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("社員總數", f"{int(data['現有社員'].sum()):,}", f"{(data['現有社員'].sum()-data['sM_total'].sum())/data['sM_total'].sum():.2%}")
        m2.metric("股金總額", f"${data['現有股金'].sum()/1e8:.2f} 億", f"{(data['現有股金'].sum()-data['sS_total'].sum())/data['sS_total'].sum():.2%}")
        m3.metric("平均收支比", f"{data['收支比'].mean():.2%}")
        m4.metric("平均逾放比", f"{data['逾放比(末)'].mean():.2%}")

        st.markdown("### 🏷️ 狀態監控")
        hr, liq, idl, std = [data[data['診斷狀態'] == s]['社名'].tolist() for s in ["🚨 高風險列管", "⚠️ 流動性緊繃", "💤 資金閒置", "✅ 穩健模範"]]
        c1, c2, c3, c4 = st.columns(4)
        def draw_card(title, names, cls):
            tags = "".join([f'<span class="name-tag">{n}</span>' for n in names]) if names else '<div style="color:#94A3B8; text-align:center; margin-top:20px;">無</div>'
            st.markdown(f'<div class="stat-card"><div class="card-header {cls}">{title}</div><div class="card-body">{tags}</div></div>', unsafe_allow_html=True)
        with c1: draw_card("🚨 高風險", hr, "header-red")
        with c2: draw_card("⚠️ 緊繃", liq, "header-orange")
        with c3: draw_card("💤 閒置", idl, "header-blue")
        with c4: draw_card("✅ 穩健", std, "header-green")

    with t2:
        st.markdown("### 🎯 全區風險分佈矩陣")
        fig_scatter = px.scatter(data, x='貸放比', y='逾放比(末)', color='診斷狀態', size='現有社員', hover_name='社名',
                                 color_discrete_map={'🚨 高風險列管': '#EF4444', '⚠️ 流動性緊繃': '#F59E0B', '💤 資金閒置': '#3B82F6', '✅ 穩健模範': '#10B981', '📊 一般狀態': '#94A3B8'},
                                 size_max=25, height=550)
        fig_scatter.add_hline(y=0.1, line_dash="dot", line_color="red", annotation_text="高風險(10%)")
        fig_scatter.update_layout(xaxis_tickformat='.1%', yaxis_tickformat='.1%', plot_bgcolor="white", legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5), margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_scatter, use_container_width=True)

    with t3:
        st.markdown("### 🏥 單一儲互社健檢")
        target_society = st.selectbox("請選擇：", data['社名'].unique())
        if target_society:
            target_data = data[data['社名'] == target_society].iloc[0]
            st.markdown(f"#### 狀態：`{target_data['診斷狀態']}`")
            metrics = ['貸放比', '儲蓄率', '逾放比(末)', '收支比', '社員成長', '股金成長']
            fig_bar = go.Figure(data=[go.Bar(name=target_society, x=metrics, y=[target_data[m] if m in target_data else target_data['社員成長率(12M)'] for m in metrics], marker_color='#3B82F6')])
            fig_bar.update_layout(height=400, yaxis_tickformat='.1%', plot_bgcolor="white", margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig_bar, use_container_width=True)

    with t4:
        st.markdown("### 📋 完整數據總表")
        st.dataframe(data.drop(columns=['sM_total', 'sS_total']).style.format({'貸放比': '{:.1%}', '逾放比(末)': '{:.2%}'}), use_container_width=True)

    with t5:
        st.markdown("### 📈 趨勢追蹤")
        selected = st.multiselect("選擇社別：", options=data['社名'].unique(), default=data['社名'].iloc[0])
        if selected:
            df_all = pd.merge(df_m, df_l[['年月', '社號', '逾放比']], on=['年月', '社號'], how='left')
            fig = px.line(df_all[df_all['社名'].isin(selected)], x='年月', y='逾放比', color='社名', markers=True)
            fig.update_layout(yaxis_tickformat='.1%', plot_bgcolor="white", legend=dict(orientation="h", y=-0.2))
            st.plotly_chart(fig, use_container_width=True)

else:
    st.markdown('<div style="text-align: center; margin-top: 50px; color: #64748B;"><h2>👋 歡迎，請載入數據開始分析</h2></div>', unsafe_allow_html=True)
