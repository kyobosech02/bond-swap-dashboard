import streamlit as st
import pandas as pd
import numpy as np
import os

# 페이지 설정 (넓은 화면 사용)
st.set_page_config(layout="wide", page_title="포지션 현황")

# --- 비밀번호 확인 기능 ---
def check_password():
    """비밀번호가 맞으면 True, 아니면 False를 반환합니다."""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    # 비밀번호 입력 창
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
    st.title("원화채권 포지션 현황")

    # 엑셀 파일 경로 (상대 경로로 수정하거나 전체 경로를 입력하세요)
    file_path = "포지션 보고양식 예시.xlsx"

    @st.cache_data
    def load_and_preprocess_data():
        # 읽어올 시트 목록
        sheets = ['9999', '9994', '9992', '9988', '7120(원화)']
        df_list = []
        
        for sheet in sheets:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet)
                df_list.append(df)
            except Exception as e:
                st.warning(f"'{sheet}' 시트를 불러오지 못했습니다. ({e})")
                
        if not df_list:
            return pd.DataFrame()
            
        data = pd.concat(df_list, ignore_index=True)
        data = data[data['펀드코드'] != 8018].copy()
        
        # 채권종류 분류 로직
        def classify_bond(row):
            ctype = str(row['채권종류']).strip()
            name = str(row['종목명']).strip()
            if ctype in ['국채', '통안채', '지방채']: return '국고/통안채'
            elif '중앙회' in name: return '특수채'
            elif ctype == '특수채': return '특수채'
            elif ctype == '금융채':
                if any(x in name for x in ['중소기업은행', '기업은행', '산업금융채권', '산금']): return '특은채'
                elif any(x in name for x in ['카드', '캐피탈']): return '여전채'
                elif '은행' in name: return '시은채'
                else: return '기타금융채'
            else: return '기타'

        data['분류_채권종류'] = data.apply(classify_bond, axis=1)
        
        # 신용등급 및 잔존만기 계산
        data['평가일자'] = pd.to_datetime(data['평가일자'])
        data['만기일자'] = pd.to_datetime(data['만기일자'])
        data['잔존년수'] = (data['만기일자'] - data['평가일자']).dt.days / 365.0
        
        def classify_maturity(years):
            if years <= 0.25: return '0.25Y'
            elif years <= 0.5: return '0.5Y'
            elif years <= 1.0: return '1Y'
            elif years <= 1.5: return '1.5Y'
            elif years <= 2.0: return '2Y'
            elif years <= 2.5: return '2.5Y'
            elif years <= 3.0: return '3Y'
            elif years <= 5.0: return '5Y'
            elif years <= 10.0: return '10Y'
            elif years <= 20.0: return '20Y'
            else: return '30Y'
            
        data['분류_잔존만기'] = data['잔존년수'].apply(classify_maturity)
        data['포지션(억원)'] = data['수량'] / 100000.0
        
        # 펀드별 카테고리 매핑
        def classify_fund(code):
            if code == 8010: return 'RP운용', '대고객 RP(8010)', '대고객 RP(8010)'
            elif code == 8013: return 'RP운용', 'CMA RP(8013)', 'CMA RP(8013)'
            elif code == 9994: return 'RP운용', '기관RP(9994)', '기관RP(9994)'
            elif code == 7120: return 'RP운용', '외화RP(7120)', '외화RP(7120)'
            elif code == 8001: return '상품채권운용', '자격', 'PD펀드(8001)'
            elif code == 8008: return '상품채권운용', '자격', '소액펀드(8008)'
            elif code == 8007: return '상품채권운용', '일반 Prop', '팀운용(8007)'
            elif code == 8016: return '상품채권운용', '일반 Prop', '부서공통(8016)'
            elif code == 8011: return '상품채권운용', '일반 Prop', 'Prop1(8011)'
            elif code == 8019: return '상품채권운용', '일반 Prop', 'Prop2(8019)'
            else: return '기타', '기타', '기타'

        data[['큰분류', '세부분류', '세세부']] = data.apply(lambda row: pd.Series(classify_fund(row['펀드코드'])), axis=1)
        return data

    df = load_and_preprocess_data()

    if not df.empty:
        # 정렬 순서 정의
        maturity_order = ['0.25Y', '0.5Y', '1Y', '1.5Y', '2Y', '2.5Y', '3Y', '5Y', '10Y', '20Y', '30Y']
        rp_order = ['대고객 RP(8010)', 'CMA RP(8013)', '기관RP(9994)', '외화RP(7120)']
        
        # --- 탭 생성 ---
        tab1, tab2 = st.tabs(["RP운용", "상품채권운용"])

        with tab1:
            st.subheader("RP운용 포지션 현황")
            rp_df = df[df['큰분류'] == 'RP운용'].copy()
            if not rp_df.empty:
                # RP 세부분류 순서 강제 지정을 위해 Categorical 사용
                rp_df['세부분류'] = pd.Categorical(rp_df['세부분류'], categories=rp_order, ordered=True)
                
                rp_pivot = pd.pivot_table(
                    rp_df, 
                    values='포지션(억원)', 
                    index=['큰분류', '세부분류', '분류_채권종류'], 
                    columns='분류_잔존만기', 
                    aggfunc='sum',
                    fill_value=0
                ).sort_index(level='세부분류') # 카테고리 순서대로 정렬
                
                # 열 순서 재배열 및 합계
                existing_cols = [col for col in maturity_order if col in rp_pivot.columns]
                rp_pivot = rp_pivot[existing_cols]
                rp_pivot['소계'] = rp_pivot.sum(axis=1)
                
                # 표 높이를 데이터 양에 맞춰 조정 (스크롤 최소화)
                table_height = (len(rp_pivot) + 1) * 35 + 40 
                st.dataframe(rp_pivot.style.format("{:,.0f}"), use_container_width=True, height=table_height)
            else:
                st.info("데이터가 없습니다.")

        with tab2:
            st.subheader("상품채권운용 포지션 현황")
            prop_df = df[df['큰분류'] == '상품채권운용']
            if not prop_df.empty:
                prop_pivot = pd.pivot_table(
                    prop_df, 
                    values='포지션(억원)', 
                    index=['큰분류', '세부분류', '세세부', '분류_채권종류'], 
                    columns='분류_잔존만기', 
                    aggfunc='sum',
                    fill_value=0
                )
                
                existing_cols = [col for col in maturity_order if col in prop_pivot.columns]
                prop_pivot = prop_pivot[existing_cols]
                prop_pivot['소계'] = prop_pivot.sum(axis=1)
                
                # 표 높이 자동 조절
                table_height = (len(prop_pivot) + 1) * 35 + 40
                st.dataframe(prop_pivot.style.format("{:,.0f}"), use_container_width=True, height=table_height)
            else:
                st.info("데이터가 없습니다.")
    else:
        st.error("불러올 데이터가 없습니다.")
