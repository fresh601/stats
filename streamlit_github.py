import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm
import os
import urllib.request
import platform # 플랫폼 확인을 위해 추가

# assemble_github 모듈은 사용자 정의 모듈이므로, 이 코드에서는 해당 모듈의 동작을 가정합니다.
# 만약 이 부분에서 문제가 발생한다면, assemble_github.py 파일을 확인해야 합니다.
try:
    from assemble_github import run_all, clean_sheet_name
except ImportError:
    st.error("🚨 'assemble_github.py' 모듈을 찾을 수 없습니다. 해당 파일을 확인하거나 경로를 설정해주세요.")
    st.stop() # 모듈이 없으면 앱 실행을 중지합니다.


# ✅ 한글 폰트 설정 (Cloud 환경 및 로컬 환경 대비)
# Streamlit Cloud 환경에서는 /tmp 디렉토리에 쓰기 권한이 있습니다.
# 로컬 환경에서는 현재 디렉토리에 저장하거나, 사용자 홈 디렉토리 등을 고려할 수 있습니다.
FONT_URL = "https://raw.githubusercontent.com/fresh601/stats/main/NanumGothic.otf"
FONT_FILE_NAME = "NanumGothic.otf"

# 플랫폼에 따라 폰트 저장 경로 설정
if platform.system() == "Windows":
    # 윈도우에서는 임시 디렉토리 사용 (혹은 사용자 지정 경로)
    FONT_PATH = os.path.join(os.getenv("TEMP"), FONT_FILE_NAME)
elif platform.system() == "Darwin": # macOS
    FONT_PATH = os.path.join("/tmp", FONT_FILE_NAME) # macOS도 /tmp 사용 가능
else: # Linux (Streamlit Cloud 포함)
    FONT_PATH = os.path.join("/tmp", FONT_FILE_NAME)

# 폰트 다운로드 (파일이 없으면 다운로드)
if not os.path.exists(FONT_PATH):
    try:
        st.info(f"폰트 다운로드 중: {FONT_URL} -> {FONT_PATH}")
        urllib.request.urlretrieve(FONT_URL, FONT_PATH)
        st.success("폰트 다운로드 완료!")
    except Exception as e:
        st.error(f"폰트 다운로드 실패: {e}")
        st.stop() # 폰트 다운로드 실패 시 앱 실행 중지

# matplotlib에 한글 폰트 등록
try:
    # 폰트 캐시를 지우고 새로고침하여 새 폰트를 인식하도록 합니다.
    fm.fontManager.addfont(FONT_PATH)
    font_name = fm.FontProperties(fname=FONT_PATH).get_name()
    plt.rcParams["font.family"] = font_name
    plt.rcParams["axes.unicode_minus"] = False # 마이너스 기호 깨짐 방지
    st.success(f"matplotlib에 '{font_name}' 폰트 적용 완료.")
except Exception as e:
    st.error(f"matplotlib 폰트 설정 실패: {e}")
    st.stop() # 폰트 설정 실패 시 앱 실행 중지


# 📌 스트림릿 기본 설정
st.set_page_config(page_title="경제지표 시각화 대시보드", layout="wide")
st.title("📊 통합 경제지표 시각화 대시보드")

# 🔄 데이터 파일명
DATA_FILE = "통합_주요지표_최종.xlsx"

# ✅ 상태 변수 초기화
if "refresh_triggered" not in st.session_state:
    st.session_state.refresh_triggered = False

# ✅ rerun 상태 변수 (Streamlit의 내장 rerun 기능 사용을 권장)
# __rerun__ 대신 st.rerun()을 직접 호출하는 것이 더 명확합니다.
# if "__rerun__" not in st.session_state:
#     st.session_state["__rerun__"] = False

# ✅ 새로고침 트리거 처리 (st.rerun()으로 대체)
# if st.session_state["__rerun__"]:
#     st.session_state["__rerun__"] = False
#     st.experimental_rerun() # 또는 st.rerun()

# ✅ 수동 수집 버튼
st.markdown("### 📂 통계 데이터 최신화")
if st.button("🌀 최신 통계 데이터 새로 수집"):
    with st.spinner("🛠 데이터 수집 중입니다..."):
        if os.path.exists(DATA_FILE):
            try:
                os.remove(DATA_FILE)
                st.info(f"기존 데이터 파일 '{DATA_FILE}' 삭제 완료.")
            except OSError as e:
                st.error(f"기존 데이터 파일 삭제 실패: {e}")
                st.stop() # 파일 삭제 실패 시 앱 중지

        run_all() # assemble_github.py의 run_all 함수 호출

        # 수집 후 파일 존재 여부 확인
        if os.path.exists(DATA_FILE):
            st.success("✅ 데이터 수집 및 파일 생성 완료! 앱을 새로고침합니다.")
            st.rerun() # 데이터 수집 후 앱 새로고침
        else:
            st.error("❗ 데이터 파일이 생성되지 않았습니다. `run_all()` 함수 확인이 필요합니다.")

# ✅ 앱 최초 로딩 시 데이터 수집
if not os.path.exists(DATA_FILE):
    with st.spinner("🔄 초기 데이터 수집 중..."):
        run_all()
    if not os.path.exists(DATA_FILE):
        st.error("❗ 초기 데이터 수집에 실패했습니다. `run_all()` 함수를 확인해주세요.")
        st.stop() # 초기 데이터 없으면 앱 실행 중지

# 📥 엑셀 파일 읽기
try:
    xls = pd.ExcelFile(DATA_FILE)
    sheet_names = xls.sheet_names
    selected_sheet = st.selectbox("📂 보고 싶은 지표를 선택하세요", sheet_names)
    df = pd.read_excel(xls, selected_sheet)
except FileNotFoundError:
    st.error(f"❗ 데이터 파일 '{DATA_FILE}'을 찾을 수 없습니다. '최신 통계 데이터 새로 수집' 버튼을 눌러주세요.")
    st.stop()
except Exception as e:
    st.error(f"엑셀 파일 읽기 오류: {e}")
    st.stop()


# ✅ 컬럼 정리
df.columns = [c.strip() for c in df.columns]

# ✅ 컬럼 후보
date_col_candidates = ["시점", "날짜", "기준년월"] # '기준년월' 추가
value_col_candidates = ["지표값", "값", "수치"] # '수치' 추가
item_col_candidates = ["항목명", "지수종류", "항목", "항목이름", "분류"] # '분류' 추가

date_col = next((col for col in date_col_candidates if col in df.columns), None)
value_col = next((col for col in value_col_candidates if col in df.columns), None)
item_col = next((col for col in item_col_candidates if col in df.columns), None)

# 필수 컬럼이 없는 경우 경고 및 중단
if not date_col or not value_col:
    st.error(f"⚠️ 필수 컬럼 ('시점' 또는 '날짜', '지표값' 또는 '값')을 찾을 수 없습니다. 현재 컬럼: {df.columns.tolist()}")
    st.stop()


# ✅ 날짜 파싱 함수
def parse_date(val):
    val = str(val).strip()
    # 분기 데이터 처리 (예: 2023Q1)
    if 'Q' in val:
        try:
            y, q = val.split('Q')
            month = {'1': '01', '2': '04', '3': '07', '4': '10'}.get(q, '01')
            return pd.to_datetime(f"{y}-{month}-01") # 일자 추가
        except ValueError:
            return pd.NaT
    # 다양한 날짜 형식 처리
    for fmt in ("%Y%m", "%Y-%m", "%Y.%m", "%Y/%m", "%Y"): # YYYY/MM 형식 추가
        try:
            return pd.to_datetime(val, format=fmt)
        except ValueError:
            continue
    return pd.NaT

# ✅ 시계열 처리 및 시각화
st.subheader(f"� {selected_sheet} 지표 분석")

df[date_col] = df[date_col].apply(parse_date)
df[value_col] = (
    df[value_col]
    .astype(str)
    .str.replace(",", "")
    .str.strip()
    .replace("", pd.NA)
)
df[value_col] = pd.to_numeric(df[value_col], errors='coerce')

# 유효한 날짜와 값만 필터링
df_cleaned = df.dropna(subset=[date_col, value_col]).sort_values(by=date_col)

# ✅ 항목 필터링
if item_col and not df_cleaned[item_col].empty:
    unique_items = df_cleaned[item_col].dropna().unique()
    if len(unique_items) > 0:
        # 기본 선택은 첫 번째 항목 또는 모든 항목
        default_selection = unique_items.tolist() if len(unique_items) <= 5 else unique_items[:1].tolist()
        selected_items = st.multiselect("🔍 항목 선택", unique_items, default=default_selection)
        filtered_df = df_cleaned[df_cleaned[item_col].isin(selected_items)]
    else:
        filtered_df = df_cleaned.copy()
        st.info("선택할 수 있는 항목이 없습니다.")
else:
    filtered_df = df_cleaned.copy()
    if item_col: # item_col이 있지만 데이터가 없는 경우
        st.info(f"'{item_col}' 컬럼에 유효한 항목 데이터가 없습니다.")
    else: # item_col 자체가 없는 경우
        st.info("항목별 필터링을 위한 '항목명' 또는 유사 컬럼이 없습니다.")

# ✅ 유효 데이터 검사
if filtered_df[date_col].notna().sum() == 0 or filtered_df[value_col].notna().sum() == 0:
    st.warning("📭 선택된 조건에 해당하는 유효한 데이터가 없습니다. 다른 지표나 항목을 선택해 보세요.")
else:
    # ✅ 시계열 그래프
    st.subheader("📈 시계열 그래프")
    fig, ax = plt.subplots(figsize=(12, 6))

    if item_col and not filtered_df[item_col].empty and len(filtered_df[item_col].unique()) > 1:
        # 여러 항목이 있을 때 hue 사용
        sns.lineplot(data=filtered_df, x=date_col, y=value_col, hue=item_col, marker='o', ax=ax)
        ax.legend(title=item_col, bbox_to_anchor=(1.05, 1), loc='upper left') # 범례 위치 조정
    else:
        # 단일 항목 또는 item_col이 없을 때
        ax.plot(filtered_df[date_col], filtered_df[value_col], marker='o', label=selected_sheet)
        ax.legend()

    ax.set_title(f"{selected_sheet} - 시계열", fontsize=16)
    ax.set_xlabel("시점", fontsize=12)
    ax.set_ylabel("값", fontsize=12)
    plt.xticks(rotation=45, ha='right') # x축 레이블 회전
    plt.tight_layout() # 레이아웃 자동 조정
    st.pyplot(fig)

    # ✅ 원시 데이터 확인
    with st.expander("🔍 원시 데이터 확인"):
        st.write("Null 값 요약:")
        st.write(filtered_df[[date_col, value_col, item_col] if item_col else [date_col, value_col]].isna().sum())
        st.write("데이터 미리보기 (상위 10개 행):")
        st.dataframe(filtered_df.head(10))

