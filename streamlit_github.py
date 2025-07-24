import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm
import os
from assemble_github import run_all, clean_sheet_name  # 사용자 정의 함수

# ✅ 한글 폰트 설정 (Windows 기준)
# 시스템 한글 폰트 자동 적용
font_paths = fm.findSystemFonts(fontpaths=["/usr/share/fonts", "/usr/local/share/fonts"])
han_fonts = [f for f in font_paths if 'Nanum' in f or 'Un' in f]

if han_fonts:
    font_path = han_fonts[0]
    font_name = fm.FontProperties(fname=font_path).get_name()
    plt.rcParams['font.family'] = font_name
    plt.rcParams['axes.unicode_minus'] = False

# 📌 스트림릿 기본 설정
st.set_page_config(page_title="경제지표 시각화 대시보드", layout="wide")
st.title("📊 통합 경제지표 시각화 대시보드")

# 🔄 데이터 로딩
DATA_FILE = "통합_주요지표_최종.xlsx"

# ✅ 상태 변수 초기화
if "refresh_triggered" not in st.session_state:
    st.session_state.refresh_triggered = False

# ✅ 사용자 수동 수집 버튼
st.markdown("### 📂 통계 데이터 최신화")

if st.button("🌀 최신 통계 데이터 새로 수집"):
    if os.path.exists(DATA_FILE):
        os.remove(DATA_FILE)
    run_all()
    st.session_state.refresh_triggered = True
    st.experimental_rerun()  # ✅ 버튼 내부에서만 호출해야 함

# ✅ 새로고침 후 상태 초기화
if st.session_state.refresh_triggered:
    st.session_state.refresh_triggered = False

# 🔁 기존 로직 유지
if not os.path.exists(DATA_FILE):
    with st.spinner("데이터 수집 중..."):
        run_all()

# 📥 엑셀 파일 읽기
xls = pd.ExcelFile(DATA_FILE)
sheet_names = xls.sheet_names
selected_sheet = st.selectbox("📂 보고 싶은 지표를 선택하세요", sheet_names)
df = pd.read_excel(xls, selected_sheet)

# ✅ 컬럼 정리
df.columns = [c.strip() for c in df.columns]

# ✅ 컬럼 후보
date_col_candidates = ["시점", "날짜"]
value_col_candidates = ["지표값", "값"]
item_col_candidates = ["항목명", "지수종류", "항목", "항목이름"]

date_col = next((col for col in date_col_candidates if col in df.columns), None)
value_col = next((col for col in value_col_candidates if col in df.columns), None)
item_col = next((col for col in item_col_candidates if col in df.columns), None)

# ✅ 날짜 변환 함수
def parse_date(val):
    val = str(val).strip()
    if 'Q' in val:
        try:
            y, q = val.split('Q')
            month = {'1': '01', '2': '04', '3': '07', '4': '10'}.get(q, '01')
            return pd.to_datetime(f"{y}-{month}")
        except:
            return pd.NaT
    for fmt in ("%Y%m", "%Y-%m", "%Y.%m", "%Y"):
        try:
            return pd.to_datetime(val, format=fmt)
        except:
            continue
    return pd.NaT

# ✅ 시계열 처리
if date_col and value_col:
    df[date_col] = df[date_col].apply(parse_date)
    df[value_col] = (
        df[value_col]
        .astype(str)
        .str.replace(",", "")
        .str.strip()
        .replace("", pd.NA)
    )
    df[value_col] = pd.to_numeric(df[value_col], errors='coerce')

    # ✅ 항목 필터링
    if item_col:
        unique_items = df[item_col].dropna().unique()
        selected_items = st.multiselect("🔍 항목 선택", unique_items, default=unique_items[:1])
        filtered_df = df[df[item_col].isin(selected_items)]
    else:
        filtered_df = df.copy()

    # ✅ 필터 후 유효성 검사
    if filtered_df[date_col].notna().sum() == 0 or filtered_df[value_col].notna().sum() == 0:
        st.warning("📭 선택된 조건에 해당하는 유효한 데이터가 없습니다.")
    else:
        # ✅ 시계열 그래프
        st.subheader("📈 시계열 그래프")
        fig, ax = plt.subplots(figsize=(12, 6))

        if item_col:
            sns.lineplot(data=filtered_df, x=date_col, y=value_col, hue=item_col, marker='o', ax=ax)
        else:
            ax.plot(filtered_df[date_col], filtered_df[value_col], marker='o')

        ax.set_title(f"{selected_sheet} - 시계열")
        ax.set_xlabel("시점")
        ax.set_ylabel("값")
        st.pyplot(fig)

    # ✅ 디버깅용 데이터 확인
    with st.expander("🔍 원시 데이터 확인"):
        st.write("Null 값 요약:")
        st.write(filtered_df[[date_col, value_col]].isna().sum())
        st.dataframe(filtered_df.head(10))
else:
    st.warning("⚠️ 시계열 그래프를 생성할 수 있는 필수 컬럼이 없습니다.")
