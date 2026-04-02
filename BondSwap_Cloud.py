import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime

# 페이지 기본 설정 (반드시 최상단에 위치)
st.set_page_config(page_title="Bond-Swap Spread 분석", page_icon="📈", layout="wide")

# ==========================================
# 0. 비밀번호 인증 로직
# ==========================================
def check_password():
    """Returns `True` if the user had the correct password."""
    def password_entered():
        # 여기에 설정한 비밀번호 입력
        if st.session_state["password"] == "kyoboh02": 
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # 보안을 위해 입력한 비밀번호 기록 삭제
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
# 1. 데이터 로드 및 전처리 함수 정의
# ==========================================
@st.cache_data
def load_data():
    file_path = "Data.xlsx"
    
    # --- 1. BOND 데이터 처리 ---
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

    # --- 2. IRS 데이터 처리 ---
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
# 메인 대시보드 실행 (비밀번호 통과 시에만 작동)
# ==========================================
if check_password():
    try:
        bond_data, irs_data = load_data()
    except Exception as e:
        st.error(f"🚨 엑셀 파일을 읽는 중 오류가 발생했습니다.\nError: {e}")
        st.stop()

    # ==========================================
    # 2. 사이드바 (UI 구조 변경: Expander 적용)
    # ==========================================
    with st.sidebar:
        # 새로 추가된 문구 (안내 메시지 박스 형태)
        st.header("⚙️ 분석 옵션 설정")
        st.markdown('<p style="font-size: 0.8rem; color: black; margin-bottom: 0;">교보증권 채권운용부 유지민 (02-3771-9160)</p>', unsafe_allow_html=True)
        
        # 첫 번째 메뉴 익스팬더
        with st.expander("기간별 Bond-Swap sp 분석", expanded=True):
            # 1. 채권 종류
            st.markdown("**1. 채권 종류**")
            bond_list = list(bond_data.keys())
            selected_bond = st.selectbox("1. 채권 종류", bond_list, index=0, label_visibility="collapsed")
            
            sample_dates = bond_data[selected_bond]['일자']
            min_date = sample_dates.min().date()
            max_date = sample_dates.max().date()
            
            # 2. 분석 기간 (기본값: 최근 1년)
            default_start_date = max_date - datetime.timedelta(days=365)
            if default_start_date < min_date:
                default_start_date = min_date
                
            st.markdown("**2. 분석 기간**")
            col_start, col_end = st.columns(2)
            with col_start:
                start_date = st.date_input("시작일", default_start_date, min_value=min_date, max_value=max_date)
            with col_end:
                end_date = st.date_input("종료일", max_date, min_value=min_date, max_value=max_date)
                
            # 3. 그래프 분석 만기
            st.markdown("**3. 그래프 분석 만기 (차트 전용)**")
            maturities_list = ['3M', '6M', '9M', '1Y', '1.5Y', '2Y', '3Y', '4Y', '5Y']
            selected_mat = st.selectbox("3. 분석 만기", maturities_list, index=3, label_visibility="collapsed")

        # 두 번째 메뉴 익스팬더 (빈칸 예약)
        with st.expander("Tab 2 옵션", expanded=False):
            st.write("-")

    # ==========================================
    # 3. 데이터 병합 및 계산 로직 (Tab 1 용)
    # ==========================================
    df_bond = bond_data[selected_bond].copy()
    bond_rename_dict = {m: f'{m}_Bond' for m in maturities_list}
    df_bond.rename(columns=bond_rename_dict, inplace=True)

    irs_dfs = []
    for m in maturities_list:
        irs_key = 'CD91' if m == '3M' else m
        temp = irs_data[irs_key].copy()
        temp.rename(columns={'금리': f'{m}_IRS'}, inplace=True)
        irs_dfs.append(temp)

    df_irs_all = irs_dfs[0]
    for temp_df in irs_dfs[1:]:
        df_irs_all = pd.merge(df_irs_all, temp_df, on='일자', how='outer')

    merged_df = pd.merge(df_bond, df_irs_all, on='일자', how='inner')

    # Spread 계산: (IRS - 채권) * 100 -> bp 단위
    for m in maturities_list:
        merged_df[f'{m}_Spread'] = (merged_df[f'{m}_IRS'] - merged_df[f'{m}_Bond']) * 100

    mask = (merged_df['일자'].dt.date >= start_date) & (merged_df['일자'].dt.date <= end_date)
    final_df = merged_df.loc[mask].reset_index(drop=True).copy()

    # ==========================================
    # 4. 메인 화면 출력 (UI 구조 변경: Tab 적용)
    # ==========================================
    st.title("📈 Bond-Swap Spread")
    st.markdown(f"**선택된 채권:** `{selected_bond}` &nbsp;|&nbsp; **분석 기간:** `{start_date} ~ {end_date}`")

    # 탭 생성
    tab1, tab2 = st.tabs(["기간별 Bond-Swap sp 분석", "Tab 2"])

    # ------------------------------------------
    # [첫 번째 탭] 기존 분석 화면
    # ------------------------------------------
    with tab1:
        if final_df.empty:
            st.warning("선택하신 기간 내에 유효한 데이터가 없습니다.")
        else:
            # --- 1) 표 출력 (최근 3개월치) ---
            st.subheader("📊 Bond-Swap 추이 (최근 3개월)")

            max_dt = final_df['일자'].max()
            three_months_ago = max_dt - pd.DateOffset(months=3)
            recent_3m_df = final_df[final_df['일자'] >= three_months_ago].sort_values('일자', ascending=False).copy()
            recent_3m_df['일자_str'] = recent_3m_df['일자'].dt.strftime('%Y-%m-%d')

            multi_columns = pd.MultiIndex.from_product(
                [maturities_list, ['채권(%)', 'IRS(%)', 'Spread(bp)']],
                names=['만기', '구분']
            )

            display_df = pd.DataFrame(index=recent_3m_df['일자_str'], columns=multi_columns)

            for m in maturities_list:
                display_df[(m, '채권(%)')] = recent_3m_df[f'{m}_Bond'].values
                display_df[(m, 'IRS(%)')] = recent_3m_df[f'{m}_IRS'].values
                display_df[(m, 'Spread(bp)')] = recent_3m_df[f'{m}_Spread'].values

            display_df.index.name = '영업일'

            def highlight_spread(s):
                return ['font-weight: 900; color: #003366; background-color: #E8F4F8;' for _ in s]

            format_dict = {}
            for m in maturities_list:
                format_dict[(m, '채권(%)')] = "{:.3f}"
                format_dict[(m, 'IRS(%)')] = "{:.3f}"
                format_dict[(m, 'Spread(bp)')] = "{:.1f}"

            idx = pd.IndexSlice
            styled_df = display_df.style.format(format_dict)\
                                        .apply(highlight_spread, subset=idx[:, idx[:, 'Spread(bp)']], axis=0)

            st.dataframe(styled_df, use_container_width=True, height=400)

            st.divider()

            # --- 2) 차트 좌우 배치 ---
            chart_col1, chart_col2 = st.columns(2)

            # [좌측 차트] 금리 및 스프레드 추이
            with chart_col1:
                st.subheader(f"📉 금리 및 스프레드 추이 ({selected_mat})")

                chart_df = final_df.sort_values('일자')
                fig1 = make_subplots(specs=[[{"secondary_y": True}]])
                irs_key_display = 'CD91' if selected_mat == '3M' else selected_mat

                fig1.add_trace(go.Scatter(x=chart_df['일자'], y=chart_df[f'{selected_mat}_Bond'], 
                                         name=f"채권금리 ({selected_bond})", line=dict(color='#2E86C1', width=2)),
                              secondary_y=False)
                fig1.add_trace(go.Scatter(x=chart_df['일자'], y=chart_df[f'{selected_mat}_IRS'], 
                                         name=f"IRS금리 ({irs_key_display})", line=dict(color='#E67E22', width=2, dash='dash')),
                              secondary_y=False)
                fig1.add_trace(go.Scatter(x=chart_df['일자'], y=chart_df[f'{selected_mat}_Spread'], 
                                         name="Spread (bp)", mode='lines',
                                         line=dict(color='rgba(108, 122, 137, 0.8)', width=1.5),
                                         fill='tozeroy', fillcolor='rgba(108, 122, 137, 0.2)'),
                              secondary_y=True)

                fig1.update_layout(height=400, hovermode="x unified", plot_bgcolor='white',
                                  legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                                  margin=dict(l=0, r=0, t=50, b=0))

                fig1.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
                fig1.update_yaxes(title_text="금리 (%)", showgrid=True, gridwidth=1, gridcolor='LightGray', secondary_y=False)
                fig1.update_yaxes(title_text="Spread (bp)", showgrid=False, secondary_y=True)

                st.plotly_chart(fig1, use_container_width=True)

                # 좌측 차트 하단 Expander
                exp_col1, exp_col2 = st.columns(2)
                with exp_col1:
                    with st.expander("📋 상세데이터"):
                        detail_df = chart_df[['일자', f'{selected_mat}_Bond', f'{selected_mat}_IRS', f'{selected_mat}_Spread']].copy()
                        detail_df['일자'] = detail_df['일자'].dt.strftime('%Y-%m-%d')
                        detail_df.columns = ['영업일', '채권(%)', 'IRS(%)', 'Spread(bp)']
                        detail_df.set_index('영업일', inplace=True)
                        st.dataframe(detail_df.style.format({'채권(%)': "{:.3f}", 'IRS(%)': "{:.3f}", 'Spread(bp)': "{:.1f}"}), use_container_width=True)

                with exp_col2:
                    with st.expander("📊 기초통계값"):
                        target_col = f'{selected_mat}_Spread'
                        stat_df = chart_df[['일자', target_col]].dropna().copy()
                        if not stat_df.empty:
                            max_idx = stat_df[target_col].idxmax()
                            min_idx = stat_df[target_col].idxmin()
                            
                            max_val = stat_df.loc[max_idx, target_col]
                            max_date = stat_df.loc[max_idx, '일자'].strftime('%Y-%m-%d')
                            min_val = stat_df.loc[min_idx, target_col]
                            min_date = stat_df.loc[min_idx, '일자'].strftime('%Y-%m-%d')
                            avg_val = stat_df[target_col].mean()
                            
                            stat_df['prev'] = stat_df[target_col].shift(1)
                            stat_df['next'] = stat_df[target_col].shift(-1)
                            local_maxs = stat_df[(stat_df[target_col] > stat_df['prev']) & (stat_df[target_col] > stat_df['next'])]
                            local_mins = stat_df[(stat_df[target_col] < stat_df['prev']) & (stat_df[target_col] < stat_df['next'])]
                            
                            prev_high_val = local_maxs.iloc[-1][target_col] if not local_maxs.empty else max_val
                            prev_high_date = local_maxs.iloc[-1]['일자'].strftime('%Y-%m-%d') if not local_maxs.empty else max_date
                            prev_low_val = local_mins.iloc[-1][target_col] if not local_mins.empty else min_val
                            prev_low_date = local_mins.iloc[-1]['일자'].strftime('%Y-%m-%d') if not local_mins.empty else min_date

                            stats_data = {
                                "항목": ["최대값", "최소값", "평균값", "전고점", "전저점"],
                                "수치(bp)": [f"{max_val:.1f}", f"{min_val:.1f}", f"{avg_val:.1f}", f"{prev_high_val:.1f}", f"{prev_low_val:.1f}"],
                                "날짜": [max_date, min_date, "-", prev_high_date, prev_low_date]
                            }
                            st.table(pd.DataFrame(stats_data).set_index("항목"))

            # [우측 차트] 만기별 Spread 커브
            with chart_col2:
                st.subheader("📉 만기별 Spread 커브 (Term Structure)")

                latest_date = final_df['일자'].max()
                latest_row = final_df[final_df['일자'] == latest_date].iloc[0]

                curve_data = {'Latest': [], 'Avg': [], 'Max': [], 'Max_Date': [], 'Min': [], 'Min_Date': []}

                for m in maturities_list:
                    col = f'{m}_Spread'
                    curve_data['Latest'].append(latest_row[col])
                    curve_data['Avg'].append(final_df[col].mean())
                    
                    max_idx = final_df[col].idxmax()
                    min_idx = final_df[col].idxmin()
                    
                    curve_data['Max'].append(final_df.loc[max_idx, col])
                    curve_data['Max_Date'].append(final_df.loc[max_idx, '일자'].strftime('%Y-%m-%d'))
                    curve_data['Min'].append(final_df.loc[min_idx, col])
                    curve_data['Min_Date'].append(final_df.loc[min_idx, '일자'].strftime('%Y-%m-%d'))

                fig2 = go.Figure()

                fig2.add_trace(go.Scatter(x=maturities_list, y=curve_data['Latest'], mode='lines+markers', 
                                          name=f"최근({latest_date.strftime('%m/%d')})", line=dict(color='red', width=2.5)))
                fig2.add_trace(go.Scatter(x=maturities_list, y=curve_data['Avg'], mode='lines+markers', 
                                          name="평균(Avg)", line=dict(color='green', dash='dash')))
                fig2.add_trace(go.Scatter(x=maturities_list, y=curve_data['Max'], mode='lines+markers', 
                                          name="최대(Max)", line=dict(color='blue', dash='dot')))
                fig2.add_trace(go.Scatter(x=maturities_list, y=curve_data['Min'], mode='lines+markers', 
                                          name="최소(Min)", line=dict(color='purple', dash='dot')))

                fig2.update_layout(height=400, hovermode="x unified", plot_bgcolor='white',
                                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                                   margin=dict(l=0, r=0, t=50, b=0))

                fig2.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
                fig2.update_yaxes(title_text="Spread (bp)", showgrid=True, gridwidth=1, gridcolor='LightGray')

                st.plotly_chart(fig2, use_container_width=True)

                # 우측 차트 하단 Expander
                with st.expander("📊 만기별 Spread 상세 데이터 및 기록 날짜"):
                    display_curve_df = pd.DataFrame(index=maturities_list)
                    display_curve_df['최근(bp)'] = [f"{val:.1f}" for val in curve_data['Latest']]
                    display_curve_df['평균(bp)'] = [f"{val:.1f}" for val in curve_data['Avg']]
                    display_curve_df['최대(bp)'] = [f"{val:.1f} ({dt})" for val, dt in zip(curve_data['Max'], curve_data['Max_Date'])]
                    display_curve_df['최소(bp)'] = [f"{val:.1f} ({dt})" for val, dt in zip(curve_data['Min'], curve_data['Min_Date'])]
                    display_curve_df.index.name = '만기'
                    
                    st.table(display_curve_df)

    # ------------------------------------------
    # [두 번째 탭] 신규 기능 예약 공간
    # ------------------------------------------
    with tab2:
        st.info("-")
