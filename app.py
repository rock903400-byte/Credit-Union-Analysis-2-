import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import uuid  # 👈 新增：用來產生安全的隨機檔名
from supabase import create_client, Client

# 1. 頁面基礎設定
st.set_page_config(page_title="儲互社決策分析中心 v7.1", layout="wide", page_icon="🏦")

# --- 初始化 Supabase 連線 ---
@st.cache_resource
def init_supabase() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase = init_supabase()
BUCKET_NAME = "excel-reports"

# --- 自定義 CSS ---
st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    .stat-card {
        background: white; border-radius: 12px; border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 1.5rem;
        height: 240px; display: flex; flex-direction: column; overflow: hidden;
    }
    .card-header {
        padding: 12px; color: white; font-weight: 700; font-size: 1.1rem;
        text-align: center; display: flex; align-items: center; justify-content: center; gap: 8px;
    }
    .header-red { background: linear-gradient(135deg, #EF4444, #991B1B); }
    .header-orange { background: linear-gradient(135deg, #F59E0B, #92400E); }
    .header-blue { background: linear-gradient(135deg, #3B82F6, #1E40AF); }
    .header-green { background: linear-gradient(135deg, #10B981, #065F46); }
    .card-body { padding: 15px; overflow-y: auto; flex-grow: 1; background: #FFFFFF; }
    .name-tag {
        display: inline-block; background: #F1F5F9; color: #1E293B; padding: 4px 12px;
        border-radius: 8px; margin: 4px; font-size: 0.9rem; border: 1px solid #CBD5E1; font-weight: 600;
    }
    .stTabs [data-baseweb="tab"] { font-size: 1.15rem; font-weight: 600; padding: 15px 20px; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- 民國年月轉換 ---
def convert_minguo_date(val):
    try:
        s = str(int(val)).strip()
        if len(s) == 5: year, month = int(s[:3]) + 1911, int(s[3:])
        elif len(s) == 4: year, month = int(s[:2]) + 1911, int(s[2:])
        else: return pd.NaT
        return pd.to_datetime(f"{year}-{month}-01")
    except: return pd.NaT

# --- 資料處理引擎 ---
@st.cache_data(show_spinner="🚀 正在啟動決策引擎與數據解析...")
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
        memGrowth = (eM-sM)/sM if sM!=0 else 0
        shrGrowth = (eS-sS)/sS if sS!=0 else 0

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

    return pd.DataFrame(rows), df_m, df_l

# --- 側邊欄與資料載入邏輯 ---
st.sidebar.markdown("## ⚙️ 決策數據中心")

query_params = st.query_params
shared_file = query_params.get("file")

data_loaded = False
data, df_m, df_l = None, None, None

# 情境 A：透過分享網址進入
if shared_file:
    st.sidebar.info(f"📁 正在從雲端載入數據...")
    try:
        res = supabase.storage.from_(BUCKET_NAME).download(shared_file)
        file_bytes = io.BytesIO(res)
        data, df_m, df_l = process_excel_only(file_bytes)
        data_loaded = True
        st.sidebar.success("✅ 雲端資料載入成功！")
    except Exception as e:
        st.sidebar.error("❌ 檔案讀取失敗，可能是連結已失效。")

# 情境 B：一般進入 (顯示上傳介面)
else:
    uploaded_file = st.sidebar.file_uploader("匯入 Excel 檔案", type=["xlsx"])
    
    if uploaded_file:
        data, df_m, df_l = process_excel_only(uploaded_file)
        data_loaded = True
        # 👈 移除撒花通知 st.toast
        
        st.sidebar.markdown("---")
        st.sidebar.markdown("### ☁️ 雲端分享功能")
        if st.sidebar.button("🚀 生成即時分享連結"):
            with st.spinner("正在安全加密並上傳至雲端..."):
                file_bytes = uploaded_file.getvalue()
                # 👈 解決方案：生成 8 碼的安全隨機檔名，避開中文字與覆蓋問題
                safe_filename = f"report_{uuid.uuid4().hex[:8]}.xlsx"
                
                try:
                    supabase.storage.from_(BUCKET_NAME).upload(
                        safe_filename,
                        file_bytes,
                        file_options={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "x-upsert": "true"}
                    )
                    
                    app_base_url = "8asdxeziyl2ozfrmkpzof3.streamlit.app" 
                    share_url = f"{app_base_url}/?file={safe_filename}"
                    
                    st.sidebar.success("✅ 上傳成功！")
                    st.sidebar.code(share_url, language="text")
                    st.sidebar.caption("複製上方連結給其他人，他們就能直接觀看此報表！")
                except Exception as e:
                    st.sidebar.error(f"上傳失敗：{e}")

# --- 主畫面渲染 ---
if data_loaded:
    t1, t2, t3, t4, t5 = st.tabs(["📊 經營總覽", "🎯 全域風險矩陣", "🏥 個社深度健檢", "📋 報表匯出", "📈 趨勢追蹤"])

    with t1:
        st.markdown("### 🏆 區會級總體指標")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("全體社員總數", f"{int(data['現有社員'].sum()):,}", f"{(data['現有社員'].sum()-data['sM_total'].sum())/data['sM_total'].sum():.2%}")
        m2.metric("全體股金總額", f"${data['現有股金'].sum()/1e8:.2f} 億", f"{(data['現有股金'].sum()-data['sS_total'].sum())/data['sS_total'].sum():.2%}")
        m3.metric("全區平均收支比", f"{data['收支比'].mean():.2%}")
        m4.metric("全區平均逾放比", f"{data['逾放比(末)'].mean():.2%}")

        st.markdown("### 🏷️ 狀態雷達監控")
        hr = data[data['診斷狀態'] == '🚨 高風險列管']['社名'].tolist()
        liq = data[data['診斷狀態'] == '⚠️ 流動性緊繃']['社名'].tolist()
        idl = data[data['診斷狀態'] == '💤 資金閒置']['社名'].tolist()
        std = data[data['診斷狀態'] == '✅ 穩健模範']['社名'].tolist()

        c1, c2, c3, c4 = st.columns(4)
        def draw_card(title, icon, names, cls):
            tags = "".join([f'<span class="name-tag">{n}</span>' for n in names]) if names else '<div style="color:#94A3B8; text-align:center; margin-top:20px;">無標的</div>'
            st.markdown(f'<div class="stat-card"><div class="card-header {cls}">{title}</div><div class="card-body">{tags}</div></div>', unsafe_allow_html=True)
        with c1: draw_card("🚨 高風險列管", "", hr, "header-red")
        with c2: draw_card("⚠️ 流動性緊繃", "", liq, "header-orange")
        with c3: draw_card("💤 資金閒置", "", idl, "header-blue")
        with c4: draw_card("✅ 穩健模範", "", std, "header-green")

    with t2:
        st.markdown("### 🎯 全區風險分佈矩陣 (散佈圖)")
        st.caption("💡 每個氣泡代表一家儲互社，氣泡越大代表社員數越多。透過十字線可快速辨識高逾放與流動性極端的社別。")
        
        fig_scatter = px.scatter(data, x='貸放比', y='逾放比(末)', 
                                 color='診斷狀態', size='現有社員', hover_name='社名',
                                 color_discrete_map={
                                     '🚨 高風險列管': '#EF4444', '⚠️ 流動性緊繃': '#F59E0B',
                                     '💤 資金閒置': '#3B82F6', '✅ 穩健模範': '#10B981', '📊 一般狀態': '#94A3B8'
                                 }, size_max=40, height=600)
        
        fig_scatter.add_hline(y=0.1, line_dash="dot", line_color="red", annotation_text="高風險紅線 (10%)")
        fig_scatter.add_vline(x=0.9, line_dash="dot", line_color="orange", annotation_text="緊繃線 (90%)")
        fig_scatter.add_vline(x=0.3, line_dash="dot", line_color="blue", annotation_text="閒置線 (30%)")
        
        fig_scatter.update_layout(xaxis_tickformat='.1%', yaxis_tickformat='.1%', plot_bgcolor="white")
        fig_scatter.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#F1F5F9')
        fig_scatter.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#F1F5F9')
        st.plotly_chart(fig_scatter, use_container_width=True)

    with t3:
        st.markdown("### 🏥 單一儲互社深度健檢報告")
        target_society = st.selectbox("請選擇要診斷的儲互社：", data['社名'].unique())
        
        if target_society:
            target_data = data[data['社名'] == target_society].iloc[0]
            global_avg = data.mean(numeric_only=True)
            
            st.markdown(f"#### 【{target_society}】目前狀態：`{target_data['診斷狀態']}`")
            
            metrics = ['貸放比', '儲蓄率', '逾放比(末)', '收支比', '社員成長率(12M)', '股金成長率(12M)']
            target_vals = [target_data[m] for m in metrics]
            avg_vals = [global_avg[m] for m in metrics]
            
            fig_bar = go.Figure(data=[
                go.Bar(name=target_society, x=metrics, y=target_vals, marker_color='#3B82F6'),
                go.Bar(name='全區平均', x=metrics, y=avg_vals, marker_color='#CBD5E1')
            ])
            fig_bar.update_layout(barmode='group', height=450, yaxis_tickformat='.1%', plot_bgcolor="white", title="關鍵指標與大盤對比")
            fig_bar.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#F1F5F9')
            st.plotly_chart(fig_bar, use_container_width=True)

    with t4:
        st.markdown("### 📋 完整診斷數據總表")
        final_table = data.drop(columns=['sM_total', 'sS_total'])
        st.dataframe(
            final_table.style.apply(lambda row: ['background-color: #FEF2F2; font-weight: bold; color: #991B1B' if '高風險' in row['診斷狀態'] else '' for _ in row], axis=1)
            .format({
                '社員成長率(12M)': '{:.2%}', '股金成長率(12M)': '{:.2%}', '貸放比': '{:.1%}', 
                '儲蓄率': '{:.2%}', '逾放比(初)': '{:.2%}', '逾放比(末)': '{:.2%}', 
                '提撥率': '{:.2%}', '收支比': '{:.2%}', '現有社員': '{:,}', '現有股金': '${:,.0f}'
            }), use_container_width=True, height=550
        )
        st.download_button("📥 一鍵匯出 AI 診斷報表 (CSV)", final_table.to_csv(index=False).encode('utf-8-sig'), "儲互社經營診斷報告.csv", "text/csv")

    with t5:
        st.markdown("### 📈 歷史趨勢對比 (含大盤基準線)")
        show_avg = st.checkbox("顯示大盤基準線 (黑色虛線)", value=True)
        
        df_all = pd.merge(df_m, df_l[['年月', '社號', '逾放比', '收支比']], on=['年月', '社號'], how='left')
        avg_df = df_all.groupby('年月').mean(numeric_only=True).reset_index()
        avg_df['社名'] = '—— 全區大盤 ——'

        selected = st.multiselect("加入比較的儲互社：", options=data['社名'].unique(), default=data['社名'].iloc[0])
        
        if selected:
            plot_data = pd.concat([df_all[df_all['社名'].isin(selected)], avg_df]) if show_avg else df_all[df_all['社名'].isin(selected)]
            
            def draw_chart(y_col, title, is_pct=True):
                fig = px.line(plot_data, x='年月', y=y_col, color='社名', title=title, markers=True, color_discrete_map={'—— 全區大盤 ——': '#1E293B'})
                fig.for_each_trace(lambda t: t.update(line=dict(dash='dash', width=3)) if t.name == '—— 全區大盤 ——' else ())
                if is_pct: fig.update_layout(yaxis_tickformat='.1%')
                fig.update_layout(hovermode="x unified", plot_bgcolor="white")
                st.plotly_chart(fig, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1: draw_chart('社員數', "👥 社員數", False)
            with c2: draw_chart('貸放比', "💰 貸放比")
            c3, c4 = st.columns(2)
            with c3: draw_chart('儲蓄率', "🏦 儲蓄率")
            with c4: draw_chart('逾放比', "⚠️ 逾放比")
            st.divider()
            draw_chart('收支比', "📈 收支比")

else:
    st.markdown("""
        <div style="text-align: center; margin-top: 100px; color: #64748B;">
            <h1 style="font-size: 3rem;">🏦 儲互社雲端決策中心</h1>
            <h2>歡迎使用</h2>
            <p>請於左側上傳 Excel 檔案，或透過分享連結直接載入最新分析數據。</p>
        </div>
    """, unsafe_allow_html=True)
