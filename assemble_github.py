import requests
import pandas as pd
from collections import defaultdict
import re
import streamlit as st

# 🔑 API 키 설정
ECOS_API_KEY = st.secrets["api"]["ECOS_API_KEY"]
INDEX_API_KEY = st.secrets["api"]["INDEX_API_KEY"]
KOSIS_API_KEY = st.secrets["api"]["KOSIS_API_KEY"]

def fetch_ecos_data():
    config = [
    {"title": "실질GDP", "stat_code": "200Y106", "period": "Q", "start": "2022Q1", "end": "2025Q2", "item_code1": "1400", "item_code2": None},
    {"title": "명목GDP", "stat_code": "200Y105", "period": "Q", "start": "2022Q1", "end": "2025Q2", "item_code1": "1400", "item_code2": None},
    {"title": "소비자물가지수", "stat_code": "901Y009", "period": "M", "start": "202401", "end": "202506", "item_code1": "0", "item_code2": None},
    {"title": "생산자물가지수(기본분류)", "stat_code": "404Y014", "period": "M", "start": "202401", "end": "202506", "item_code1": "*AA", "item_code2": None},
    {"title": "수출물가지수(기본분류)", "stat_code": "402Y014", "period": "M", "start": "202401", "end": "202506", "item_code1": "*AA", "item_code2": "W"},
    {"title": "수입물가지수(기본분류)", "stat_code": "401Y015", "period": "M", "start": "202401", "end": "202506", "item_code1": "*AA", "item_code2": "W"},
    {"title": "환율(달러)", "stat_code": "731Y006", "period": "M", "start": "202401", "end": "202506", "item_code1": "0000002", "item_code2": "0000100"},
    {"title": "환율(위안)", "stat_code": "731Y006", "period": "M", "start": "202401", "end": "202506", "item_code1": "0000007", "item_code2": "0000100"},
    {"title": "환율(엔화)", "stat_code": "731Y006", "period": "M", "start": "202401", "end": "202506", "item_code1": "0000006", "item_code2": "0000100"},
    {"title": "선행종합지수", "stat_code": "901Y067", "period": "M", "start": "202401", "end": "202506", "item_code1": "I16A", "item_code2": None},
    {"title": "동행종합지수", "stat_code": "901Y067", "period": "M", "start": "202401", "end": "202506", "item_code1": "I16B", "item_code2": None},
    {"title": "후행종합지수", "stat_code": "901Y067", "period": "M", "start": "202401", "end": "202506", "item_code1": "I16C", "item_code2": None},
    {"title": "선행지수순환변동치", "stat_code": "901Y067", "period": "M", "start": "202401", "end": "202506", "item_code1": "I16E", "item_code2": None},
    {"title": "동행지수순환변동치", "stat_code": "901Y067", "period": "M", "start": "202401", "end": "202506", "item_code1": "I16D", "item_code2": None},
    {"title": "향후경기전망CSI", "stat_code": "511Y002", "period": "M", "start": "202401", "end": "202506", "item_code1": "FMBB", "item_code2": "99988"},
    {"title": "소비자심리지수", "stat_code": "511Y002", "period": "M", "start": "202401", "end": "202506", "item_code1": "FME", "item_code2": "99988"},
    {"title": "기업경기실사지수(실적)_전체", "stat_code": "512Y013", "period": "M", "start": "202401", "end": "202506", "item_code1": "99988", "item_code2": "AX"},
    {"title": "기업경기실사지수(실적)_제조업", "stat_code": "512Y013", "period": "M", "start": "202401", "end": "202506", "item_code1": "C0000", "item_code2": "AX"},
    {"title": "기업경기실사지수(실적)_비제조업", "stat_code": "512Y013", "period": "M", "start": "202401", "end": "202506", "item_code1": "Y9900", "item_code2": "AX"},
    {"title": "기업경기실사지수(전망)_전체", "stat_code": "512Y014", "period": "M", "start": "202401", "end": "202506", "item_code1": "99988", "item_code2": "BX"},
    {"title": "기업경기실사지수(전망)_제조업", "stat_code": "512Y014", "period": "M", "start": "202401", "end": "202506", "item_code1": "C0000", "item_code2": "BX"},
    {"title": "기업경기실사지수(전망)_비제조업", "stat_code": "512Y014", "period": "M", "start": "202401", "end": "202506", "item_code1": "Y9900", "item_code2": "BX"},
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
