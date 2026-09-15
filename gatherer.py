import os
import json
import time
import requests
import re
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from deep_translator import GoogleTranslator

# 1. 깃허브 금고(Secrets)에서 열쇠 꺼내오기
gcp_secret_json = os.environ.get("GCP_SERVICE_ACCOUNT")
serp_api_key = os.environ.get("SERPAPI_KEY")

# 2. 구글 시트 연결
creds_dict = json.loads(gcp_secret_json)
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
gc = gspread.authorize(credentials)

queue_sheet = gc.open("zenifix_DB").worksheet("수집예약")
result_sheet = gc.open("zenifix_DB").worksheet("수집결과")

now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
print(f"[{now_str}] 🤖 수집 전용 로봇 가동 시작...")

# 언어 매핑 사전
COUNTRY_LANG_MAP = {
    "USA": "en", "UK": "en", "Australia": "en", "Canada": "en",
    "Germany": "de", "Austria": "de", "France": "fr",
    "Japan": "ja", "Vietnam": "vi", "Thailand": "th",
    "Spain": "es", "Mexico": "es", "UAE": "ar", "Italy": "it",
    "China": "zh-CN", "Taiwan": "zh-TW", "Russia": "ru",
    "Brazil": "pt", "Indonesia": "id", "Poland": "pl"
}

# 3. 대기열(Queue) 확인
records = queue_sheet.get_all_records()
pending_tasks = []
for idx, row in enumerate(records, start=2): # 엑셀은 2번째 줄부터
    status = str(row.get("상태", ""))
    reserve_time = str(row.get("예약일시", ""))
    # 상태가 대기중이고, 지정한 시간이 현재 시간과 같거나 지났을 경우 색출
    if status == "대기중" and reserve_time <= now_str:
        pending_tasks.append((idx, row))

if not pending_tasks:
    print("✅ 현재 시간이 된 수집 예약 건이 없습니다. 로봇을 종료합니다.")
    exit()

print(f"🔍 총 {len(pending_tasks)}건의 지역 수집 예약을 처리합니다.")

def extract_emails(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers, timeout=5)
        soup = BeautifulSoup(response.text, 'html.parser')
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        found_emails = re.findall(email_pattern, soup.get_text(separator=' '))
        for link in soup.find_all('a'):
            href = link.get('href')
            if href and href.startswith('mailto:'):
                found_emails.append(href.replace('mailto:', '').split('?')[0])
        valid_emails = {e.lower() for e in found_emails if not e.endswith(('.png', '.jpg', '.gif'))}
        return list(valid_emails)
    except:
        return []

# 4. 본격적인 수집 및 번역 실행
for row_num, row_data in pending_tasks:
    target_loc = str(row_data.get("타깃국가", ""))
    keyword = str(row_data.get("검색키워드", ""))
    page_count = int(row_data.get("페이지수", 2))

    print(f"➡️ [{target_loc}] 지역 데이터 수집 시작...")

    # 입력된 텍스트(예: Paris, France)를 보고 언어 코드 찾기
    lang_code = "en"
    for country, code in COUNTRY_LANG_MAP.items():
        if country.lower() in target_loc.lower():
            lang_code = code
            break

    try:
        if lang_code != "en":
            translated_keyword = GoogleTranslator(source='auto', target=lang_code).translate(keyword)
        else:
            translated_keyword = keyword
    except:
        translated_keyword = keyword

    search_queries = [f"{keyword} {target_loc}"]
    if translated_keyword != keyword:
        search_queries.append(f"{translated_keyword} {target_loc}")

    collected_count = 0
    api_limit_hit = False
    
    for query in search_queries:
        if api_limit_hit: break
        for page in range(page_count):
            offset = page * 10
            api_url = f"https://serpapi.com/search.json?engine=google&q={query}&start={offset}&api_key={serp_api_key}"
            try:
                res = requests.get(api_url).json()
                if 'error' in res:
                    print(f"🚨 API 에러: {res['error']}")
                    api_limit_hit = True
                    break
                
                if 'organic_results' in res:
                    for item in res['organic_results']:
                        company = item.get('title', '이름 없음')
                        url = item.get('link', '')
                        if url.endswith('.pdf'): continue
                        
                        emails = extract_emails(url)
                        if emails:
                            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            # 수집결과 시트에 1줄씩 저장: [수집일시, 국가명, 업체명, 웹사이트, 이메일, 검색키워드]
                            result_sheet.append_row([current_time, target_loc, company, url, emails[0], keyword])
                            collected_count += 1
                    time.sleep(1) # 차단 방지 휴식
            except Exception as e:
                print(f"검색 중 오류 발생: {e}")

    # 예약 시트 상태 업데이트 (대기중 -> 수집완료)
    queue_sheet.update_cell(row_num, 5, "수집완료")
    queue_sheet.update_cell(row_num, 6, str(collected_count)) # 수집된 건수 기록
    print(f"✅ [{target_loc}] 수집 완료! 총 {collected_count}건 구글 시트에 저장됨.")
    
print("🎉 예약된 수집 작업이 모두 완료되었습니다!")
