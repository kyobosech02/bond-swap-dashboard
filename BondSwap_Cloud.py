import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime

# 페이지 기본 설정
st.set_page_config(page_title="Bond-Swap Spread 분석", page_icon="📈", layout="wide")

# ==========================================
# 0. 비밀번호 인증 로직
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == "kyoboh02":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("비밀번호가 틀렸습니다.")
        return False
    return True

# ==========================================
# 1. 데이터 로드 및 전처리
# ==========================================
@st.cache_data
def load_data():
    file_path = "Data.xlsx"
   
    # BOND 데이터
    bond_raw = pd.read_excel(file_path, sheet_name='BOND', header=None)
    bond_dict = {}
    maturities = ['3M', '6M', '9M', '1Y', '1.5Y', '2Y', '3Y', '4Y', '5Y']
   
    for i in range(0, bond_raw.shape[1], 10):
        bond_name = bond_raw.iloc[1, i+1]
        if pd.isna(bond_name):
            bond_name = bond_raw.iloc[2, i]
            if pd.isna(bond_name): continue
           
        df_temp = bond_raw.iloc[4:, i:i+10].copy()
        df_temp.columns = ['일자'] + maturities
        df_temp['일자'] = pd.to_datetime(df_temp['일자'], errors='coerce')
        for col in maturities:
            df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce')
           
        bond_dict[bond_name] = df_temp.dropna(subset=['일자']).sort_values('일자').reset_index(drop=True)

    # IRS 데이터
    irs_raw = pd.read_excel(file_path, sheet_name='IRS', header=None)
    irs_dict = {}
    irs_mat_names = ['CD91', '6M', '9M', '1Y', '1.5Y', '2Y', '3Y', '4Y', '5Y']
   
    for i, mat_label in enumerate(irs_mat_names):
        col_idx = i * 2
        df_temp = irs_raw.iloc[4:, col_idx:col_idx+2].copy()
        df_temp.columns = ['일자', '금리']
        df_temp['일자'] = pd.to_datetime(df_temp['일자'], errors='coerce')
        df_temp['금리'] = pd.to_numeric(df_temp['금리'], errors='coerce')
       
        irs_dict[mat_label] = df_temp.dropna(subset=['일자']).sort_values('일자').reset_index(drop=True)
       
    return bond_dict, irs_dict

# ==========================================
# 메인 대시보드 실행
# ==========================================
if check_password():
    try:
        bond_data, irs_data = load_data()
    except Exception as e:
        st.error(f"🚨 엑셀 파일을 읽는 중 오류가 발생했습니다.\nError: {e}")
        st.stop()

    maturities_list = ['3M', '6M', '9M', '1Y', '1.5Y', '2Y', '3Y', '4Y', '5Y']
    x_numeric = [0.25, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]

    with st.sidebar:
        st.header("⚙️ 분석 옵션 설정")
        st.caption("교보증권 채권운용부 유지민 (02-3771-9160)")
       
        with st.expander("1) Bond Swap spread 분석", expanded=True):
            bond_list = list(bond_data.keys())
            selected_bond = st.selectbox("1. 채권 종류", bond_list, index=0)
            sample_dates = bond_data[selected_bond]['일자']
            min_date, max_date = sample_dates.min().date(), sample_dates.max().date()
            start_date = st.date_input("시작일", max_date - datetime.timedelta(days=365), min_value=min_date, max_value=max_date, key='t1_sd')
            end_date = st.date_input("종료일", max_date, min_value=min_date, max_value=max_date, key='t1_ed')
            selected_mat = st.selectbox("3. 분석 만기", maturities_list, index=3)

        with st.expander("2) Credit Spread 분석", expanded=False):
            t2_bond1 = st.selectbox("채권 종류 1 (비교)", bond_list, index=bond_list.index('은행채 AAA') if '은행채 AAA' in bond_list else 0)
            t2_bond2 = st.selectbox("채권 종류 2 (기준)", bond_list, index=bond_list.index('국고채권') if '국고채권' in bond_list else 0)
            t2_start_date = st.date_input("시작일 ", max_date - datetime.timedelta(days=365), key='t2_sd')
            t2_end_date = st.date_input("종료일 ", max_date, key='t2_ed')
            t2_selected_mat = st.selectbox("분석 만기 ", maturities_list, index=5, key='t2_mat')

    # [데이터 처리 로직 - Tab 1 & 2 동일하게 유지]
    # (중략된 부분은 기존 데이터 병합 로직과 동일)
    df_bond = bond_data[selected_bond].copy()
    irs_dfs = [irs_data['CD91' if m == '3M' else m].copy().rename(columns={'금리': f'{m}_IRS'}) for m in maturities_list]
    df_irs_all = irs_dfs[0]
    for temp_df in irs_dfs[1:]: df_irs_all = pd.merge(df_irs_all, temp_df, on='일자', how='outer')
    merged_df = pd.merge(df_bond.rename(columns={m: f'{m}_Bond' for m in maturities_list}), df_irs_all, on='일자', how='inner')
    for m in maturities_list: merged_df[f'{m}_Spread'] = (merged_df[f'{m}_IRS'] - merged_df[f'{m}_Bond']) * 100
    final_df = merged_df[(merged_df['일자'].dt.date >= start_date) & (merged_df['일자'].dt.date <= end_date)].sort_values('일자')

    df_t2_b1 = bond_data[t2_bond1].copy().rename(columns={m: f'{m}_B1' for m in maturities_list})
    df_t2_b2 = bond_data[t2_bond2].copy().rename(columns={m: f'{m}_B2' for m in maturities_list})
    df_t2_merged = pd.merge(df_t2_b1, df_t2_b2, on='일자', how='inner')
    for m in maturities_list: df_t2_merged[f'{m}_Spread'] = (df_t2_merged[f'{m}_B1'] - df_t2_merged[f'{m}_B2']) * 100
    t2_final_df = df_t2_merged[(df_t2_merged['일자'].dt.date >= t2_start_date) & (df_t2_merged['일자'].dt.date <= t2_end_date)].sort_values('일자')

    st.title("📈 Bond-Swap Spread Dashboard")
    tab1, tab2 = st.tabs(["1) Bond Swap Spread 분석", "2) Credit Spread 분석"])

    # --- [탭 1] ---
    with tab1:
        if not final_df.empty:
            st.subheader("📊 Bond-Swap 추이 및 커브 분석")
            c1, c2 = st.columns(2)
            with c1:
                # (기존 왼쪽 차트 fig1 로직 유지)
                chart_df_t1 = final_df[['일자', f'{selected_mat}_Bond', f'{selected_mat}_IRS', f'{selected_mat}_Spread']].dropna()
                fig1 = make_subplots(specs=[[{"secondary_y": True}]])
                fig1.add_trace(go.Scatter(x=chart_df_t1['일자'], y=chart_df_t1[f'{selected_mat}_Bond'], name="채권금리"), secondary_y=False)
                fig1.add_trace(go.Scatter(x=chart_df_t1['일자'], y=chart_df_t1[f'{selected_mat}_IRS'], name="IRS금리", line=dict(dash='dash')), secondary_y=False)
                fig1.add_trace(go.Scatter(x=chart_df_t1['일자'], y=chart_df_t1[f'{selected_mat}_Spread'], name="Spread(bp)", fill='tozeroy'), secondary_y=True)
                fig1.update_layout(height=400, hovermode="x unified", plot_bgcolor='white', margin=dict(l=0,r=0,t=30,b=0))
                st.plotly_chart(fig1, use_container_width=True)

            with c2:
                latest_date = final_df['일자'].max()
                latest_row = final_df[final_df['일자'] == latest_date].iloc[0]
                
                # 우측 차트: 만기별 커브 (수정됨)
                fig2 = make_subplots(specs=[[{"secondary_y": True}]])
                
                # 금리 데이터 (초기 숨김)
                fig2.add_trace(go.Scatter(x=x_numeric, y=[latest_row[f'{m}_Bond'] for m in maturities_list], 
                                          mode='lines+markers', name=f"{selected_bond} 금리(%)", 
                                          line=dict(color='#2E86C1'), visible='legendonly'), secondary_y=False)
                fig2.add_trace(go.Scatter(x=x_numeric, y=[latest_row[f'{m}_IRS'] for m in maturities_list], 
                                          mode='lines+markers', name="IRS 금리(%)", 
                                          line=dict(color='#E67E22', dash='dash'), visible='legendonly'), secondary_y=False)
                
                # 스프레드 데이터 (항상 표시)
                fig2.add_trace(go.Scatter(x=x_numeric, y=[latest_row[f'{m}_Spread'] for m in maturities_list], 
                                          mode='lines+markers', name="최근 Spread(bp)", line=dict(color='red', width=3)), secondary_y=True)
                fig2.add_trace(go.Scatter(x=x_numeric, y=[final_df[f'{m}_Spread'].mean() for m in maturities_list], 
                                          mode='lines+markers', name="평균 Spread", line=dict(color='green', dash='dot')), secondary_y=True)

                fig2.update_layout(height=400, hovermode="x unified", plot_bgcolor='white', legend=dict(orientation="h", y=1.1), margin=dict(l=0,r=0,t=30,b=0))
                fig2.update_xaxes(tickvals=x_numeric, ticktext=maturities_list, showgrid=True, gridcolor='LightGray')
                fig2.update_yaxes(title_text="금리 (%)", secondary_y=False, showgrid=True, gridcolor='LightGray')
                fig2.update_yaxes(title_text="Spread (bp)", secondary_y=True, showgrid=False)
                st.plotly_chart(fig2, use_container_width=True)

    # --- [탭 2] ---
    with tab2:
        if not t2_final_df.empty:
            st.subheader("📊 크레딧 스프레드 추이 및 커브 분석")
            t2_c1, t2_c2 = st.columns(2)
            with t2_c1:
                # (기존 탭2 왼쪽 차트 fig_t2_lt 로직 유지)
                chart_df_t2 = t2_final_df[['일자', f'{t2_selected_mat}_B1', f'{t2_selected_mat}_B2', f'{t2_selected_mat}_Spread']].dropna()
                fig_t2_lt = make_subplots(specs=[[{"secondary_y": True}]])
                fig_t2_lt.add_trace(go.Scatter(x=chart_df_t2['일자'], y=chart_df_t2[f'{t2_selected_mat}_B1'], name=t2_bond1), secondary_y=False)
                fig_t2_lt.add_trace(go.Scatter(x=chart_df_t2['일자'], y=chart_df_t2[f'{t2_selected_mat}_B2'], name=t2_bond2, line=dict(dash='dash')), secondary_y=False)
                fig_t2_lt.add_trace(go.Scatter(x=chart_df_t2['일자'], y=chart_df_t2[f'{t2_selected_mat}_Spread'], name="Spread(bp)", fill='tozeroy'), secondary_y=True)
                fig_t2_lt.update_layout(height=400, hovermode="x unified", plot_bgcolor='white', margin=dict(l=0,r=0,t=30,b=0))
                st.plotly_chart(fig_t2_lt, use_container_width=True)

            with t2_c2:
                latest_t2 = t2_final_df.iloc[-1]
                
                # 우측 차트: 만기별 커브 (수정됨 - 탭1과 동일한 선그래프 형식)
                fig_t2_rt = make_subplots(specs=[[{"secondary_y": True}]])
                
                # 금리 데이터 (초기 숨김 설정)
                fig_t2_rt.add_trace(go.Scatter(x=x_numeric, y=[latest_t2[f'{m}_B1'] for m in maturities_list], 
                                             mode='lines+markers', name=f"{t2_bond1} 금리(%)", 
                                             line=dict(color='#2E86C1'), visible='legendonly'), secondary_y=False)
                fig_t2_rt.add_trace(go.Scatter(x=x_numeric, y=[latest_t2[f'{m}_B2'] for m in maturities_list], 
                                             mode='lines+markers', name=f"{t2_bond2} 금리(%)", 
                                             line=dict(color='#E67E22', dash='dash'), visible='legendonly'), secondary_y=False)
                
                # 스프레드 데이터 (선형으로 변경)
                fig_t2_rt.add_trace(go.Scatter(x=x_numeric, y=[latest_t2[f'{m}_Spread'] for m in maturities_list], 
                                             mode='lines+markers', name="최근 Spread(bp)", 
                                             line=dict(color='red', width=3)), secondary_y=True)
                fig_t2_rt.add_trace(go.Scatter(x=x_numeric, y=[t2_final_df[f'{m}_Spread'].mean() for m in maturities_list], 
                                             mode='lines+markers', name="평균 Spread", 
                                             line=dict(color='green', dash='dot')), secondary_y=True)
                
                # Max/Min 포인트를 선에 추가
                fig_t2_rt.add_trace(go.Scatter(x=x_numeric, y=[t2_final_df[f'{m}_Spread'].max() for m in maturities_list], 
                                             mode='markers', name="최대 Spread", marker=dict(symbol='triangle-up', color='blue')), secondary_y=True)
                fig_t2_rt.add_trace(go.Scatter(x=x_numeric, y=[t2_final_df[f'{m}_Spread'].min() for m in maturities_list], 
                                             mode='markers', name="최소 Spread", marker=dict(symbol='triangle-down', color='purple')), secondary_y=True)

                fig_t2_rt.update_layout(height=400, hovermode="x unified", plot_bgcolor='white', legend=dict(orientation="h", y=1.1), margin=dict(l=0,r=0,t=30,b=0))
                fig_t2_rt.update_xaxes(tickvals=x_numeric, ticktext=maturities_list, showgrid=True, gridcolor='LightGray')
                fig_t2_rt.update_yaxes(title_text="금리 (%)", secondary_y=False, showgrid=True, gridcolor='LightGray')
                fig_t2_rt.update_yaxes(title_text="Spread (bp)", secondary_y=True, showgrid=False)
                st.plotly_chart(fig_t2_rt, use_container_width=True)
