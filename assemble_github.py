import requests
import pandas as pd
from collections import defaultdict
import re
import streamlit as st
from datetime import datetime, timedelta

# 🔑 API 키 설정
ECOS_API_KEY = st.secrets["api"]["ECOS_API_KEY"]
INDEX_API_KEY = st.secrets["api"]["INDEX_API_KEY"]
KOSIS_API_KEY = st.secrets["api"]["KOSIS_API_KEY"]

# === 날짜 계산 유틸 ===
def get_current_quarter():
    """오늘 날짜 기준 현재 분기 문자열 반환 (YYYYQn)"""
    now = datetime.today()
    quarter = (now.month - 1) // 3 + 1
    return f"{now.year}Q{quarter}"

def get_quarter_from_date(date):
    """날짜 객체를 YYYYQn 형식으로 변환"""
    quarter = (date.month - 1) // 3 + 1
    return f"{date.year}Q{quarter}"

def get_month_str(date):
    """날짜 객체를 YYYYMM 형식으로 변환"""
    return date.strftime("%Y%m")

# === 자동 기간 생성 ===
def get_auto_period(years_back=3):
    """현재 날짜 기준 start, end 자동 계산"""
    now = datetime.today()

    # 시작 시점
    start_date = now.replace(year=now.year - years_back)
    start_quarter = get_quarter_from_date(start_date)
    start_month = get_month_str(start_date)

    # 종료 시점
    end_quarter = get_current_quarter()
    end_month = get_month_str(now)

    return start_quarter, end_quarter, start_month, end_month

# === API 호출 ===
def fetch_ecos_data():
    start_q, end_q, start_m, end_m = get_auto_period(years_back=3)

    config = [
        {"title": "실질GDP", "stat_code": "200Y106", "period": "Q", "start": start_q, "end": end_q, "item_code1": "1400", "item_code2": None},
        {"title": "명목GDP", "stat_code": "200Y105", "period": "Q", "start": start_q, "end": end_q, "item_code1": "1400", "item_code2": None},
        {"title": "소비자물가지수", "stat_code": "901Y009", "period": "M", "start": start_m, "end": end_m, "item_code1": "0", "item_code2": None},
        {"title": "생산자물가지수(기본분류)", "stat_code": "404Y014", "period": "M", "start": start_m, "end": end_m, "item_code1": "*AA", "item_code2": None},
        {"title": "수출물가지수(기본분류)", "stat_code": "402Y014", "period": "M", "start": start_m, "end": end_m, "item_code1": "*AA", "item_code2": "W"},
        {"title": "수입물가지수(기본분류)", "stat_code": "401Y015", "period": "M", "start": start_m, "end": end_m, "item_code1": "*AA", "item_code2": "W"},
        {"title": "환율(달러)", "stat_code": "731Y006", "period": "M", "start": start_m, "end": end_m, "item_code1": "0000002", "item_code2": "0000100"},
        {"title": "환율(위안)", "stat_code": "731Y006", "period": "M", "start": start_m, "end": end_m, "item_code1": "0000007", "item_code2": "0000100"},
        {"title": "환율(엔화)", "stat_code": "731Y006", "period": "M", "start": start_m, "end": end_m, "item_code1": "0000006", "item_code2": "0000100"},
        {"title": "선행종합지수", "stat_code": "901Y067", "period": "M", "start": start_m, "end": end_m, "item_code1": "I16A", "item_code2": None},
        {"title": "동행종합지수", "stat_code": "901Y067", "period": "M", "start": start_m, "end": end_m, "item_code1": "I16B", "item_code2": None},
        {"title": "후행종합지수", "stat_code": "901Y067", "period": "M", "start": start_m, "end": end_m, "item_code1": "I16C", "item_code2": None},
    ]
    result = {}

    for c in config:
        if c["item_code2"]:
            url = f"https://ecos.bok.or.kr/api/StatisticSearch/{ECOS_API_KEY}/json/kr/1/100/{c['stat_code']}/{c['period']}/{c['start']}/{c['end']}/{c['item_code1']}/{c['item_code2']}"
        else:
            url = f"https://ecos.bok.or.kr/api/StatisticSearch/{ECOS_API_KEY}/json/kr/1/100/{c['stat_code']}/{c['period']}/{c['start']}/{c['end']}/{c['item_code1']}"

        try:
            response = requests.get(url)
            items = response.json().get('StatisticSearch', {}).get('row', [])
        except:
            continue

        if not items:
            continue

        df = pd.DataFrame([{
            "항목명": item.get('ITEM_NAME1', ''),
            "단위": item.get('UNIT_NAME', ''),
            "시점": item.get('TIME', ''),
            "지표값": item.get('DATA_VALUE', '')
        } for item in items])
        result[f"[ECOS] {c['title']}"] = df
    return result



def fetch_index_go_data():
    config = [
        {"title": "국내총생산 및 경제성장률", "ix_code": "2736", "stats_code": "273601", "period": "2024:2025"},
        {"title": "시장금리", "ix_code": "1073", "stats_code": "107301", "period": "2024:2025"},
        {"title": "국제수지", "ix_code": "2735", "stats_code": "273501", "period": "2024:2025"},
    ]
    result = {}

    for c in config:
        url = "https://www.index.go.kr/unity/openApi/sttsJsonViewer.do"
        params = {
            'idntfcId': INDEX_API_KEY,
            'statsCode': c['stats_code'],
            'ixCode': c['ix_code'],
            'period': c['period']
        }

        try:
            response = requests.get(url, params=params)
            items = response.json()
        except:
            continue

        if not items:
            continue

        df = pd.DataFrame([{
            '항목이름': item.get('항목이름', ''),
            '시점': item.get('시점', ''),
            '값': item.get('값', ''),
            '단위': item.get('단위', '')
        } for item in items])
        result[f"[지표누리] {c['title']}"] = df
    return result


def fetch_kosis_data():
    url = f"https://kosis.kr/openapi/Param/statisticsParameterData.do?method=getList&apiKey={KOSIS_API_KEY}&itmId=T02+T03+T04+&objL1=0+1+2+4+3+&format=json&jsonVD=Y&prdSe=M&newEstPrdCnt=12&orgId=101&tblId=DT_1J22042"
    result = {}

    try:
        response = requests.get(url)
        items = response.json()
    except:
        return result

    if not items:
        return result

    df = pd.DataFrame([{
        '통계명': item.get('TBL_NM', ''),
        '지수종류': item.get('C1_NM', ''),
        '항목': item.get('ITM_NM', ''),
        '날짜': item.get('PRD_DE', ''),
        '단위': item.get('UNIT_NM', ''),
        '값': item.get('DT', '')
    } for item in items])

    title = items[0].get('TBL_NM', 'KOSIS_지표')
    result[f"[KOSIS] {title}"] = df
    return result

def clean_sheet_name(name):
    # 엑셀 시트 이름에 사용할 수 없는 문자 제거
    return re.sub(r'[:\\/*?\[\]]', '', name)[:31]  # 최대 31자 제한

def save_to_excel(data_frames: dict, filename="통합_주요지표_최종.xlsx"):
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for title, df in data_frames.items():
            sheet_name = clean_sheet_name(title)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"[완료] {sheet_name} 저장됨.")
    print(f"✅ 모든 데이터 저장 완료 → '{filename}'")

def run_all():
    all_data = {}
    all_data.update(fetch_ecos_data())
    all_data.update(fetch_index_go_data())
    all_data.update(fetch_kosis_data())
    save_to_excel(all_data)


# 🔃 실행
if __name__ == "__main__":
    run_all()
