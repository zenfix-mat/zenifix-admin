import streamlit as st
import pandas as pd
import requests
import re
import time
import smtplib
import streamlit.components.v1 as components
import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from deep_translator import GoogleTranslator

# --- 구글 시트(DB) 라이브러리 ---
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# 웹페이지 기본 설정
st.set_page_config(page_title="zenifix Global Admin", page_icon="🚀", layout="wide")
st.title("🚀 zenifix Global B2B Admin Dashboard")

# ==========================================
# [DB 연동] 구글 스프레드시트 초기 설정
# ==========================================
db_connected = False
blacklist_emails = [] 

try:
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    gc = gspread.authorize(credentials)
    
    # 1. 메인 발송 이력 시트
    db_sheet = gc.open("zenifix_DB").sheet1
    db_connected = True
    
    # 2. 수신거부 탭
    try:
        blacklist_sheet = gc.open("zenifix_DB").worksheet("수신거부")
        blacklist_emails = blacklist_sheet.col_values(1) 
    except:
        blacklist_sheet = gc.open("zenifix_DB").add_worksheet(title="수신거부", rows="1000", cols="2")
        blacklist_sheet.update_cell(1, 1, "이메일")
        blacklist_sheet.update_cell(1, 2, "등록일시")

    # 3. 발송예약 탭
    try:
        queue_sheet = gc.open("zenifix_DB").worksheet("발송예약")
    except:
        queue_sheet = gc.open("zenifix_DB").add_worksheet(title="발송예약", rows="1000", cols="9")
        headers = ["예약일", "이메일", "국가명", "웹사이트", "제목", "본문", "상태", "타깃유형", "등록일시"]
        for i, h in enumerate(headers, 1):
            queue_sheet.update_cell(1, i, h)

    # 4. 템플릿관리 탭 (영구 저장소)
    try:
        template_sheet = gc.open("zenifix_DB").worksheet("템플릿관리")
        template_records = template_sheet.get_all_values()
    except:
        template_sheet = gc.open("zenifix_DB").add_worksheet(title="템플릿관리", rows="100", cols="4")
        template_sheet.append_row(["타깃유형", "언어", "제목", "본문"])
        template_records = [["타깃유형", "언어", "제목", "본문"]]

except Exception as e:
    st.sidebar.error(f"구글 DB 연결 실패: {e}")
    st.sidebar.warning("발송 이력 및 템플릿 저장이 작동하지 않을 수 있습니다.")


# ==========================================
# [영구 템플릿] 구글 시트에서 템플릿 불러오기
# ==========================================
if 'email_templates' not in st.session_state:
    templates = {}
    
    # DB에 저장된 템플릿이 있으면 불러오기
    if db_connected and len(template_records) > 1:
        for row in template_records[1:]:
            if len(row) >= 4:
                tgt, lng, sub, bdy = row[0], row[1], row[2], row[3]
                if tgt not in templates:
                    templates[tgt] = {}
                templates[tgt][lng] = {"subject": sub, "body": bdy}
    else:
        # DB가 비어있으면 기본 템플릿 세팅 및 DB에 최초 기록
        default_body = """<p>Dear Cosmetics Purchasing Team,</p>
<p>I hope this email finds you well.</p>
<p>I am writing from zenifix, a premium K-Beauty skincare brand based in Seoul. We would like to politely request your team's review of zenifix products for a potential retail partnership in your market.</p>
<p>We offer 14 core SKUs across two highly effective collections—our Noni Line (7 SKUs) and Ginkgo Line (7 SKUs). What truly sets zenifix apart is our exceptional ingredient concentration. Our formulations feature natural Noni and Ginkgo extracts <strong>ranging from 21% to 58% (210,000 ppm – 580,000 ppm)</strong> depending on the SKU. We differentiate our products through this uncompromising raw material content rather than generic marketing claims.</p>
<p>zenifix 브랜드를 참고하실 수 있도록 아래에 간단한 이미지를 첨부하였습니다:</p>
<p><img src="https://zenifix.net/img/zenifix_BrandDeck_main.png" alt="zenifix Brand Overview" style="max-width: 800px; width: 100%; height: auto;"></p>
<p>To explore our complete Brand Deck, including full product details, current global sales channels, and our active SNS presence, please visit our official website:<br>
👉 <strong>Official Brand Deck: <a href="https://zenifix.net">https://zenifix.net</a></strong></p>
<p>If your team finds our brand suitable for your market after the initial review, please reply to this email. We would be happy to discuss further possibilities and details.</p>
<p>Thank you for your time and consideration.</p>
<p>Best regards,<br>
Global Partnership Team<br>
zenifix<br>
zenifix@wellsfnd.com</p>
<p><small><i>*If you do not wish to receive further emails, please reply with 'Unsubscribe'.</i></small></p>"""
        
        templates = {
            "바이어 (유통/입점)": {
                "English": {
                    "subject": "[Partnership Proposal] Premium K-Beauty: 580,000ppm Skincare by zenifix",
                    "body": default_body
                }
            }
        }
        # 최초 기본 템플릿을 구글 시트에 자동 기록
        if db_connected:
            template_sheet.append_row(["바이어 (유통/입점)", "English", templates["바이어 (유통/입점)"]["English"]["subject"], default_body])

    st.session_state.email_templates = templates


# 탭 분리
tab1, tab2, tab3 = st.tabs(["📥 1. 이메일 수집 (Gathering)", "📧 2. 자동 발송 (Sending)", "📊 3. 데이터 대시보드 (통계)"])

# ==========================================
# [탭 1] 글로벌 이메일 수집 (국가/도시 자유 직접 입력 & 듀얼 검색)
# ==========================================
with tab1:
    st.header("글로벌 이메일 자동 수집기")
    
    # 💡 수집된 엑셀 데이터를 보관할 '기억 장치' 초기화
    if 'gathered_files' not in st.session_state:
        st.session_state.gathered_files = {}

    # 번역 지원 국가 사전 (도메인 또는 지역 이름에 이 단어가 포함되면 해당 언어로 번역)
    COUNTRY_LANG_MAP = {
        "USA": "en", "UK": "en", "Australia": "en", "Canada": "en",
        "Ireland": "en", "New Zealand": "en", "India": "en", "Philippines": "en",
        "Germany": "de", "Austria": "de", "France": "fr", 
        "Japan": "ja", "Vietnam": "vi", "Thailand": "th", 
        "Spain": "es", "Mexico": "es", "UAE": "ar", "Italy": "it",
        "Korea": "ko", "China": "zh-CN", "Taiwan": "zh-TW", "Russia": "ru",
        "Brazil": "pt", "Indonesia": "id", "Poland": "pl"
    }
    
    col1, col2 = st.columns(2)
    with col1:
        serp_api_key = st.text_input("SerpApi Key (필수)", type="password")
        # 🎯 선택창(Multiselect)에서 쉼표로 구분하는 '자유 텍스트 입력창'으로 확장 복구!
        locations_input = st.text_input("타깃 국가 및 도시 (쉼표로 구분하여 자유롭게 복수 입력)", value="USA, New York, London, UK")
        # 쉼표(,)를 기준으로 텍스트를 쪼개서 리스트로 자동 변환합니다.
        selected_locations = [loc.strip() for loc in locations_input.split(",") if loc.strip()]
        
    with col2:
        search_keyword = st.text_input("검색 키워드 (영문+현지어 듀얼 검색됨)", value="korean cosmetics distributor contact")
        page_count = st.number_input("검색어당 페이지 수", min_value=1, max_value=10, value=2)

    if st.button("🔍 이메일 수집 시작", type="primary"):
        if not serp_api_key or not selected_locations:
            st.error("SerpApi Key와 타깃 국가/도시를 최소 1개 이상 입력해 주세요!")
        else:
            # 새로운 수집 시작 시 기억 장치 비우기
            st.session_state.gathered_files = {}
            
            with st.spinner("다국어 번역 및 이메일 수집 중입니다... (입력된 지역이 많을수록 시간이 소요됩니다)"):
                for location in selected_locations:
                    st.markdown(f"### 🌍 {location} 수집 현황")
                    status_text = st.empty()
                    location_results = []
                    
                    # 입력된 지역 이름(예: Paris, France) 안에 사전의 국가명이 포함되어 있는지 확인하여 언어 코드 추출
                    lang_code = "en"
                    for country, code in COUNTRY_LANG_MAP.items():
                        if country.lower() in location.lower():
                            lang_code = code
                            break
                    
                    # 현지어 자동 번역
                    try:
                        if lang_code != "en":
                            from deep_translator import GoogleTranslator
                            translated_keyword = GoogleTranslator(source='auto', target=lang_code).translate(search_keyword)
                        else:
                            translated_keyword = search_keyword
                    except:
                        translated_keyword = search_keyword 
                    
                    # 영어 원본과 현지어 번역본 모두 검색 리스트에 담기
                    search_queries = [f"{search_keyword} {location}"]
                    if translated_keyword != search_keyword:
                        search_queries.append(f"{translated_keyword} {location}")

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

                    api_limit_hit = False
                    for query in search_queries:
                        if api_limit_hit: break
                        
                        status_text.info(f"🔍 검색 중: {query} ...")
                        for page in range(page_count):
                            offset = page * 10
                            api_url = f"https://serpapi.com/search.json?engine=google&q={query}&start={offset}&api_key={serp_api_key}"
                            try:
                                response = requests.get(api_url).json()
                                
                                if 'error' in response:
                                    st.error(f"🚨 API 에러 발생: {response['error']}")
                                    api_limit_hit = True
                                    break
                                    
                                if 'organic_results' in response:
                                    for item in response['organic_results']:
                                        company_name = item.get('title', '이름 없음')
                                        website_url = item.get('link', '')
                                        if website_url.endswith('.pdf'): continue
                                        
                                        emails = extract_emails(website_url)
                                        if emails:
                                            # 발송 탭(탭 2)과의 연동을 위해 컬럼명은 '국가명'으로 유지하되, 내용은 '지역(location)'이 들어갑니다.
                                            location_results.append({
                                                "업체명": company_name, "국가명": location, "웹사이트": website_url,
                                                "담당자(유추)": "Cosmetics Purchasing Team", "이메일": emails[0], "수집상태": "대기중"
                                            })
                                        time.sleep(1) 
                            except Exception as e:
                                status_text.warning(f"검색 중 일시적 오류 발생: {e}")
                    
                    # 💡 수집 완료된 데이터를 기억 장치(session_state)에 저장
                    if location_results:
                        df = pd.DataFrame(location_results)
                        df = df.drop_duplicates(subset=['이메일'], keep='first')
                        
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False)
                        
                        st.session_state.gathered_files[location] = {
                            "count": len(df),
                            "data": output.getvalue()
                        }
                        status_text.success(f"🎉 {location} 수집 및 저장 완료! (순수 이메일 {len(df)}건)")
                    else:
                        if not api_limit_hit:
                            status_text.warning(f"⚠️ {location}에서 수집된 이메일이 없습니다.")
                    
                    st.divider()

    # 💡 기억 장치에 저장된 엑셀 파일 다운로드 버튼 노출 (새로고침 방어)
    if st.session_state.gathered_files:
        st.subheader("📥 수집 완료된 지역별 엑셀 다운로드")
        for location, file_info in st.session_state.gathered_files.items():
            # 파일명에 들어갈 수 없는 특수기호(쉼표 등)를 안전하게 치환
            safe_loc_name = location.replace(' ', '_').replace(',', '')
            st.download_button(
                label=f"📥 {location} 엑셀 파일 다운로드 ({file_info['count']}건)", 
                data=file_info['data'],
                file_name=f"Zenifix_Buyers_{safe_loc_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_btn_{safe_loc_name}"
            )


# ==========================================
# [탭 2] 글로벌 이메일 자동 발송기
# ==========================================
with tab2:
    st.header("글로벌 이메일 자동 발송기")
    
    if db_connected:
        st.success("✅ 구글 스프레드시트 연동 완료! 템플릿 및 발송 이력이 자동 관리됩니다.")

    st.markdown("""
    <style>
    button[kind="primary"] { background-color: #03C75A !important; border-color: #03C75A !important; color: white !important; }
    button[kind="primary"]:hover { background-color: #028a3f !important; border-color: #028a3f !important; }
    </style>
    """, unsafe_allow_html=True)

    # --- 템플릿 및 발송 옵션 설정 ---
    st.subheader("🎯 템플릿 선택 및 관리")
    
    with st.expander("➕ 새로운 타깃 그룹 및 언어 템플릿 추가하기"):
        new_target = st.text_input("새로운 타깃 그룹 이름 (예: VIP 바이어)")
        new_lang = st.text_input("새로운 발송 언어 (예: Spanish)")
        if st.button("템플릿 목록에 추가"):
            if new_target and new_lang:
                if new_target not in st.session_state.email_templates:
                    st.session_state.email_templates[new_target] = {}
                new_subject = f"[{new_target}] Partnership with zenifix"
                new_body = f"<p>Dear {new_target} Team,</p>\n<p>내용을 입력하세요.</p>"
                
                st.session_state.email_templates[new_target][new_lang] = {"subject": new_subject, "body": new_body}
                
                # DB에도 새 템플릿 추가
                if db_connected:
                    template_sheet.append_row([new_target, new_lang, new_subject, new_body])
                
                st.success(f"'{new_target}' - '{new_lang}' 추가 완료!")
                st.rerun()

    col3, col4 = st.columns(2)
    with col3:
        target_type = st.selectbox("1. 타깃 그룹을 선택하세요", list(st.session_state.email_templates.keys()))
    with col4:
        available_languages = list(st.session_state.email_templates[target_type].keys())
        selected_language = st.selectbox("2. 발송 언어를 선택하세요", available_languages)

    st.divider()

    # --- 이메일 편집기 ---
    st.subheader("📝 이메일 미리보기 및 직접 편집")
    st.info("💡 내용 수정 후 아래의 '영구 저장' 버튼을 누르시면 다음에 접속해도 이 내용이 그대로 유지됩니다.")
    
    current_subject = st.session_state.email_templates[target_type][selected_language]["subject"]
    current_body = st.session_state.email_templates[target_type][selected_language]["body"]
    
    edited_subject = st.text_input("📝 이메일 제목 (수정 가능)", value=current_subject)
    edited_html_body = st.text_area("🔧 이메일 본문 (HTML 태그 통째로 자유 수정)", value=current_body, height=350)

    # 👇 [핵심 기능] 템플릿 영구 저장 버튼
    if st.button("💾 현재 수정한 제목과 본문을 '현재 템플릿'으로 영구 저장", type="primary", use_container_width=True):
        # 1. 세션 스테이트(화면) 업데이트
        st.session_state.email_templates[target_type][selected_language]["subject"] = edited_subject
        st.session_state.email_templates[target_type][selected_language]["body"] = edited_html_body
        
        # 2. 구글 시트(DB) 업데이트
        if db_connected:
            try:
                records = template_sheet.get_all_values()
                found_row_idx = -1
                for i, row in enumerate(records):
                    if i > 0 and row[0] == target_type and row[1] == selected_language:
                        found_row_idx = i + 1 # gspread는 1번부터 인덱스 시작
                        break
                
                if found_row_idx != -1:
                    template_sheet.update_cell(found_row_idx, 3, edited_subject)
                    template_sheet.update_cell(found_row_idx, 4, edited_html_body)
                else:
                    template_sheet.append_row([target_type, selected_language, edited_subject, edited_html_body])
                
                st.toast("🎉 템플릿이 구글 시트에 영구 저장되었습니다!")
            except Exception as e:
                st.error(f"DB 저장 중 오류: {e}")
        else:
            st.warning("DB 연결이 끊어져 임시로만 저장되었습니다.")

    st.markdown("##### 👁️ 실제 수신자가 받아볼 이메일 미리보기")
    with st.container(border=True):
        st.markdown(f"**제목:** {edited_subject}")
        st.divider()
        components.html(edited_html_body, height=400, scrolling=True)

    st.divider()

    # --- 관리자 계정 설정 및 발송 스케줄링 (하단 배치) ---
    st.subheader("⚙️ 관리자 계정 설정 및 발송 스케줄링")
    st.info("이메일 발송 권한 및 스케줄링을 설정하는 보안 영역입니다.")
    
    with st.container(border=True):
        col_admin1, col_admin2 = st.columns(2)
        with col_admin1:
            login_email = st.text_input("개인 로그인 이메일 (예: zeni@wellsfnd.com)")
            app_password = st.text_input("16자리 앱 비밀번호", type="password")
        with col_admin2:
            sender_email = st.text_input("발송자 이메일 (From: 공통메일)", value="zenifix@wellsfnd.com")
            
        st.markdown("##### 🛡️ 수신거부(Unsubscribe) 자동 동기화")
        if st.button("🔄 내 수신함에서 '수신거부' 메일 찾아 DB에 자동 저장하기"):
            if not login_email or not app_password:
                st.warning("로그인 이메일과 앱 비밀번호를 먼저 입력해 주세요.")
            else:
                with st.spinner("수신함을 스캔하여 'Unsubscribe' 답장을 찾고 있습니다..."):
                    try:
                        mail = imaplib.IMAP4_SSL("imap.gmail.com")
                        mail.login(login_email, app_password)
                        mail.select("inbox")
                        
                        status, messages = mail.search(None, 'BODY "Unsubscribe"')
                        email_ids = messages[0].split()
                        
                        new_unsubs = 0
                        for e_id in email_ids:
                            res, msg_data = mail.fetch(e_id, '(RFC822)')
                            for response_part in msg_data:
                                if isinstance(response_part, tuple):
                                    msg = email.message_from_bytes(response_part[1])
                                    from_header = msg.get("From")
                                    email_match = re.search(r'<(.+?)>', str(from_header))
                                    sender = email_match.group(1) if email_match else str(from_header)
                                    
                                    if sender not in blacklist_emails and "mailer-daemon" not in sender.lower():
                                        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        blacklist_sheet.append_row([sender, current_time])
                                        blacklist_emails.append(sender)
                                        new_unsubs += 1
                                        
                        mail.logout()
                        st.toast(f"✅ 동기화 완료! {new_unsubs}명의 새로운 수신거부 이메일이 DB에 저장되었습니다.")
                    except Exception as e:
                        st.error(f"동기화 중 오류 발생: {e}")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 발송 예약 입력 영역
    col_date, col_file = st.columns(2)
    with col_date:
        scheduled_date = st.date_input("📅 달력에서 예약 발송 일자를 선택하세요", min_value=datetime.today().date())
    with col_file:
        uploaded_file = st.file_uploader("📥 수집한 바이어 '엑셀 파일'을 올려주세요.", type=["xlsx"])
    
    delay_seconds = st.slider("메일 발송 간격 조절 (즉시 발송 시 적용, 단위: 초)", min_value=10, max_value=300, value=180, step=10)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 발송 컨트롤 버튼들
    col_btn1, col_btn2, col_btn3 = st.columns(3)
    
    with col_btn1:
        if st.button("🧪 내 메일로 테스트 1건 발송", use_container_width=True):
            if not login_email or not app_password:
                st.warning("로그인 이메일과 앱 비밀번호를 먼저 입력해 주세요.")
            else:
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(login_email, app_password)
                    msg = MIMEMultipart()
                    msg['From'] = f"zenifix Team <{sender_email}>"
                    msg['To'] = login_email  
                    msg['Subject'] = edited_subject
                    final_html = f"<html><body>{edited_html_body}</body></html>"
                    msg.attach(MIMEText(final_html, 'html'))
                    server.send_message(msg)
                    server.quit()
                    st.toast("✅ 테스트 메일이 성공적으로 발송되었습니다!")
                except Exception as e:
                    st.error(f"테스트 발송 실패: {e}")

    with col_btn2:
        if st.button("⚡ 즉시 대량 발송 시작 (예약 안 함)", use_container_width=True):
            if not uploaded_file or not login_email or not app_password:
                st.error("엑셀 파일, 로그인 이메일, 앱 비밀번호를 모두 입력해 주세요!")
            else:
                df = pd.read_excel(uploaded_file)
                st.info(f"총 {len(df)}명의 대상에게 즉시 발송을 시작합니다...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(login_email, app_password)
                    
                    success_count = 0
                    skip_count = 0 
                    
                    for index, row in df.iterrows():
                        buyer_email = str(row.get('이메일', '')).strip()
                        buyer_country = str(row.get('국가명', '미확인'))
                        buyer_website = str(row.get('웹사이트', '미확인'))
                        
                        if buyer_email in blacklist_emails:
                            status_text.text(f"🚫 수신거부 대상 제외됨: {buyer_email}")
                            skip_count += 1
                            progress_bar.progress((index + 1) / len(df))
                            continue
                        
                        msg = MIMEMultipart()
                        msg['From'] = f"zenifix Team <{sender_email}>"
                        msg['To'] = buyer_email
                        msg['Subject'] = edited_subject
                        msg.add_header('reply-to', sender_email)
                        
                        final_html = f"<html><body>{edited_html_body}</body></html>"
                        msg.attach(MIMEText(final_html, 'html'))
                        
                        try:
                            server.send_message(msg)
                            success_count += 1
                            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            if db_connected:
                                try:
                                    db_sheet.append_row([
                                        current_time, buyer_email, buyer_country, 
                                        buyer_website, target_type, selected_language, "성공"
                                    ])
                                except Exception as e:
                                    pass
                                    
                            status_text.text(f"✅ 발송 완료 ({buyer_country}): {buyer_email}")
                        except:
                            status_text.text(f"❌ 발송 실패: {buyer_email}")
                        
                        progress_bar.progress((index + 1) / len(df))
                        if index < len(df) - 1:
                            status_text.text(f"⏳ 스팸 방지를 위해 {delay_seconds}초 대기 중...")
                            time.sleep(delay_seconds)
                            
                    server.quit()
                    st.success(f"🎉 총 {success_count}건 실시간 발송 완료! (수신거부 스킵: {skip_count}명)")
                    
                except Exception as e:
                    st.error(f"🚨 이메일 로그인 실패. 오류: {e}")

    with col_btn3:
        if st.button("📅 지정한 날짜로 예약 등록", type="primary", use_container_width=True):
            if not uploaded_file:
                st.error("엑셀 파일을 먼저 올려주세요!")
            elif not db_connected:
                st.error("구글 DB와 연결되지 않아 예약을 등록할 수 없습니다.")
            else:
                df = pd.read_excel(uploaded_file)
                st.info(f"총 {len(df)}명의 대상을 {scheduled_date} 예약 대기열에 등록합니다...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                success_count = 0
                skip_count = 0 
                
                try:
                    existing_queue_emails = queue_sheet.col_values(2) 
                except:
                    existing_queue_emails = []

                for index, row in df.iterrows():
                    buyer_email = str(row.get('이메일', '')).strip()
                    buyer_country = str(row.get('국가명', '미확인'))
                    buyer_website = str(row.get('웹사이트', '미확인'))
                    
                    if buyer_email in blacklist_emails:
                        status_text.text(f"🚫 수신거부 대상 제외됨: {buyer_email}")
                        skip_count += 1
                        progress_bar.progress((index + 1) / len(df))
                        continue
                        
                    if buyer_email in existing_queue_emails:
                        status_text.text(f"⚠️ 이미 예약된 바이어 제외됨: {buyer_email}")
                        skip_count += 1
                        progress_bar.progress((index + 1) / len(df))
                        continue
                    
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    final_html = f"<html><body>{edited_html_body}</body></html>"
                    
                    try:
                        queue_sheet.append_row([
                            str(scheduled_date), buyer_email, buyer_country, buyer_website, 
                            edited_subject, final_html, "대기중", target_type, current_time
                        ])
                        success_count += 1
                        status_text.text(f"✅ 예약 등록 완료 ({buyer_country}): {buyer_email}")
                    except Exception as e:
                        status_text.text(f"❌ DB 기록 실패: {e}")
                    
                    progress_bar.progress((index + 1) / len(df))
                    time.sleep(1.5) 
                        
                st.success(f"🎉 총 {success_count}건 예약 완료! (수신거부/중복 제외: {skip_count}명)")

    # --- 예약 현황 모니터링 대시보드 ---
    st.divider()
    st.subheader("📋 현재 발송 예약 대기열 (Queue) 현황")
    
    if db_connected:
        try:
            queue_records = queue_sheet.get_all_records()
            if queue_records:
                df_queue = pd.DataFrame(queue_records)
                if '상태' in df_queue.columns:
                    df_pending = df_queue[df_queue['상태'] == '대기중']
                else:
                    df_pending = pd.DataFrame()
                
                if not df_pending.empty:
                    total_pending = len(df_pending)
                    st.info(f"💡 현재 총 **{total_pending}건**의 메일이 발송 대기 중입니다.")
                    
                    queue_summary = df_pending.groupby('예약일').size().reset_index(name='발송 예정 건수')
                    
                    col_q1, col_q2 = st.columns([1, 2])
                    with col_q1:
                        st.markdown("**📅 날짜별 예약 요약**")
                        st.dataframe(queue_summary, hide_index=True, use_container_width=True)
                        
                    with col_q2:
                        st.markdown("**🔍 세부 예약 리스트 (최근 등록순)**")
                        display_cols = ['예약일', '국가명', '이메일', '타깃유형']
                        valid_cols = [col for col in display_cols if col in df_pending.columns]
                        st.dataframe(df_pending[valid_cols].iloc[::-1], hide_index=True, use_container_width=True)
                else:
                    st.success("🎉 현재 대기 중인 발송 예약이 없습니다.")
            else:
                st.markdown("아직 등록된 예약 데이터가 없습니다.")
        except Exception as e:
            st.warning("대기열 정보를 불러오는 중입니다... (데이터가 비어있거나 새로고침이 필요합니다)")


# ==========================================
# [탭 3] 글로벌 발송 통계 대시보드
# ==========================================
with tab3:
    import plotly.express as px
    
    st.header("📊 글로벌 발송 데이터 대시보드")
    
    if not db_connected:
        st.warning("🚨 구글 스프레드시트와 연결되어 있지 않아 데이터를 불러올 수 없습니다.")
    else:
        if st.button("🔄 최신 데이터 불러오기", use_container_width=True):
            st.rerun()
            
        st.divider()
        
        try:
            raw_data = db_sheet.get_all_values()
            if len(raw_data) > 1:
                df_stats = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                
                total_sent = len(df_stats)
                total_unsubs = len(blacklist_emails) if 'blacklist_emails' in locals() else 0
                total_countries = df_stats['국가명'].nunique() if '국가명' in df_stats.columns else 0
                
                col1, col2, col3 = st.columns(3)
                col1.metric(label="🚀 총 발송 성공", value=f"{total_sent} 건")
                col2.metric(label="🌍 도달 국가 수", value=f"{total_countries} 개국")
                col3.metric(label="🚫 수신 거부 (블랙리스트)", value=f"{total_unsubs} 건")
                
                st.divider()
                
                col_chart1, col_chart2 = st.columns(2)
                
                with col_chart1:
                    st.subheader("📍 국가별 발송 비중")
                    if '국가명' in df_stats.columns:
                        country_counts = df_stats['국가명'].value_counts().reset_index()
                        country_counts.columns = ['국가명', '발송건수']
                        fig_pie = px.pie(country_counts, values='발송건수', names='국가명', hole=0.4, 
                                         color_discrete_sequence=px.colors.sequential.Teal)
                        st.plotly_chart(fig_pie, use_container_width=True)
                        
                with col_chart2:
                    st.subheader("🎯 타깃 그룹별 발송 현황")
                    if '타깃유형' in df_stats.columns:
                        target_counts = df_stats['타깃유형'].value_counts().reset_index()
                        target_counts.columns = ['타깃유형', '발송건수']
                        fig_bar = px.bar(target_counts, x='타깃유형', y='발송건수', text_auto=True,
                                         color='타깃유형', color_discrete_sequence=px.colors.qualitative.Pastel)
                        st.plotly_chart(fig_bar, use_container_width=True)

                st.divider()
                
                st.subheader("📝 최근 발송 이력 (최신 100건)")
                st.dataframe(df_stats.iloc[::-1].head(100), use_container_width=True)
                
            else:
                st.info("💡 아직 구글 시트에 기록된 발송 데이터가 없습니다. 첫 콜드 메일을 발송하시면 통계가 자동으로 생성됩니다.")
                
        except Exception as e:
            st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
