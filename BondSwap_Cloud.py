import streamlit as st
import pandas as pd
import numpy as np

# 페이지 설정
st.set_page_config(layout="wide", page_title="포지션 및 델타 현황")

# --- 비밀번호 확인 기능 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if st.session_state["password_correct"]:
        return True

    st.title("보안 인증")
    password = st.text_input("비밀번호를 입력하세요.", type="password")
    if st.button("로그인"):
        if password == "123456":
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False

if check_password():
    st.title("원화채권 포지션 및 델타 현황")

    # 엑셀 파일 경로
    file_path = "포지션 보고양식 예시.xlsx"

    @st.cache_data
    def load_and_preprocess_data():
        sheets = ['9999', '9994', '9992', '9988', '7120(원화)']
        df_list = []
        for sheet in sheets:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet)
                df_list.append(df)
            except Exception as e:
                st.warning(f"'{sheet}' 시트를 불러오지 못했습니다. ({e})")
        
        if not df_list: return pd.DataFrame()
        data = pd.concat(df_list, ignore_index=True)
        data = data[data['펀드코드'] != 8018].copy()
        
        # 1. 채권종류 분류
        def classify_bond(row):
            ctype, name = str(row['채권종류']).strip(), str(row['종목명']).strip()
            if ctype in ['국채', '통안채', '지방채']: return '국고/통안채'
            elif '중앙회' in name or ctype == '특수채': return '특수채'
            elif ctype == '금융채':
                if any(x in name for x in ['중소기업은행', '기업은행', '산업금융채권', '산금']): return '특은채'
                elif any(x in name for x in ['카드', '캐피탈']): return '여전채'
                elif '은행' in name: return '시은채'
                else: return '기타금융채'
            return '기타'

        data['분류_채권종류'] = data.apply(classify_bond, axis=1)
        
        # 2. 잔존만기 분류
        data['평가일자'] = pd.to_datetime(data['평가일자'])
        data['만기일자'] = pd.to_datetime(data['만기일자'])
        data['잔존년수'] = (data['만기일자'] - data['평가일자']).dt.days / 365.0
        
        def classify_maturity(years):
            bins = [0, 0.25, 0.5, 1, 1.5, 2, 2.5, 3, 5, 10, 20, 100]
            labels = ['0.25Y', '0.5Y', '1Y', '1.5Y', '2Y', '2.5Y', '3Y', '5Y', '10Y', '20Y', '30Y']
            for i, b in enumerate(bins[1:]):
                if years <= b: return labels[i]
            return '30Y'
        data['분류_잔존만기'] = data['잔존년수'].apply(classify_maturity)
        
        # 3. 포지션(억원) 및 델타 계산
        # 포지션 = 수량 / 100,000
        data['포지션'] = data['수량'] / 100000.0
        # 델타 = (평가금액 * 수정듀레이션) / 100,000
        data['델타'] = (data['평가금액'] * data['수정듀레이션']) / 100000.0
        
        # 4. 펀드별 카테고리 매핑
        def classify_fund(code):
            if code == 8010: return 'RP운용', '대고객 RP(8010)', '대고객 RP(8010)'
            elif code == 8013: return 'RP운용', 'CMA RP(8013)', 'CMA RP(8013)'
            elif code == 9994: return 'RP운용', '기관RP(9994)', '기관RP(9994)'
            elif code == 7120: return 'RP운용', '외화RP(7120)', '외화RP(7120)'
            elif code in [8001, 8008]: return '상품채권운용', '자격', 'PD(8001)' if code==8001 else '소액(8008)'
            elif code in [8007, 8016, 8011, 8019]: return '상품채권운용', '일반 Prop', \
                {8007:'팀운용', 8016:'부서공통', 8011:'Prop1', 8019:'Prop2'}[code]
            return '기타', '기타', '기타'

        data[['큰분류', '세부분류', '세세부']] = data.apply(lambda row: pd.Series(classify_fund(row['펀드코드'])), axis=1)
        return data

    # 스타일 적용 함수: 소수점 제거 반영 ({:,.0f})
    def style_dataframe(df):
        return df.style.format("{:,.0f}") \
            .set_properties(**{'border': '1px solid black', 'text-align': 'center'}) \
            .set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#e0e0e0'), ('border', '2px solid black'), ('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('border', '1px solid #444')]}
            ])

    # 데이터 로드
    df = load_and_preprocess_data()

    if not df.empty:
        # 상단 선택 UI
        st.write("### 🔍 보기 설정")
        view_type = st.radio("표시할 데이터를 선택하세요:", ["포지션", "델타"], horizontal=True)
        
        # 선택된 데이터에 따라 집계 대상 컬럼 설정
        target_col = "포지션" if view_type == "포지션" else "델타"
        
        maturity_order = ['0.25Y', '0.5Y', '1Y', '1.5Y', '2Y', '2.5Y', '3Y', '5Y', '10Y', '20Y', '30Y']
        rp_order = ['대고객 RP(8010)', 'CMA RP(8013)', '기관RP(9994)', '외화RP(7120)']
        
        tab1, tab2 = st.tabs([f"📊 RP운용 ({view_type})", f"📈 상품채권운용 ({view_type})"])

        with tab1:
            rp_df = df[df['큰분류'] == 'RP운용'].copy()
            if not rp_df.empty:
                rp_df['세부분류'] = pd.Categorical(rp_df['세부분류'], categories=rp_order, ordered=True)
                rp_pivot = pd.pivot_table(rp_df, values=target_col, 
                                         index=['큰분류', '세부분류', '분류_채권종류'], 
                                         columns='분류_잔존만기', aggfunc='sum', fill_value=0).sort_index(level='세부분류')
                
                # 열 순서 맞추기 및 합계 계산
                cols = [c for c in maturity_order if c in rp_pivot.columns]
                rp_pivot = rp_pivot[cols]
                rp_pivot['합계'] = rp_pivot.sum(axis=1) 
                
                # 높이 계산 및 출력
                h = (len(rp_pivot) + 1) * 38 + 50
                st.dataframe(style_dataframe(rp_pivot), use_container_width=True, height=int(h))
            else:
                st.info("데이터가 없습니다.")

        with tab2:
            prop_df = df[df['큰분류'] == '상품채권운용']
            if not prop_df.empty:
                prop_pivot = pd.pivot_table(prop_df, values=target_col, 
                                           index=['큰분류', '세부분류', '세세부', '분류_채권종류'], 
                                           columns='분류_잔존만기', aggfunc='sum', fill_value=0)
                
                cols = [c for c in maturity_order if c in prop_pivot.columns]
                prop_pivot = prop_pivot[cols]
                prop_pivot['합계'] = prop_pivot.sum(axis=1)
                
                h = (len(prop_pivot) + 1) * 38 + 50
                st.dataframe(style_dataframe(prop_pivot), use_container_width=True, height=int(h))
            else:
                st.info("데이터가 없습니다.")
    else:
        st.error("데이터 로드 실패")
