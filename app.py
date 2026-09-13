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
# [DB 연동] 구글 스프레드시트 초기 설정 (+ 수신거부 & 예약 탭 관리)
# ==========================================
db_connected = False
blacklist_emails = [] # 수신거부 리스트 보관용

try:
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    gc = gspread.authorize(credentials)
    
    # 1. 메인 발송 이력 시트
    db_sheet = gc.open("zenifix_DB").sheet1
    db_connected = True
    
    # 2. '수신거부' 탭 세팅
    try:
        blacklist_sheet = gc.open("zenifix_DB").worksheet("수신거부")
        blacklist_emails = blacklist_sheet.col_values(1) 
    except:
        blacklist_sheet = gc.open("zenifix_DB").add_worksheet(title="수신거부", rows="1000", cols="2")
        blacklist_sheet.update_cell(1, 1, "이메일")
        blacklist_sheet.update_cell(1, 2, "수신거부일시")

    # 3. [새로 추가됨] 24시간 '발송예약(Queue)' 탭 세팅
    try:
        queue_sheet = gc.open("zenifix_DB").worksheet("발송예약")
    except:
        queue_sheet = gc.open("zenifix_DB").add_worksheet(title="발송예약", rows="1000", cols="9")
        headers = ["예약일", "이메일", "국가명", "웹사이트", "제목", "본문", "상태", "타깃유형", "등록일시"]
        for i, h in enumerate(headers, 1):
            queue_sheet.update_cell(1, i, h)
            
except Exception as e:
    st.sidebar.error(f"구글 DB 연결 실패: {e}")

# ==========================================
# [영구 템플릿] 제목과 내용을 세트로 관리
# ==========================================
if 'email_templates' not in st.session_state:
    st.session_state.email_templates = {
        "바이어 (유통/입점)": {
            "English": {
                "subject": "[Partnership Proposal] Premium K-Beauty: 580,000ppm High-Content Skincare by zenifix",
                "body": """<p>Dear Cosmetics Purchasing Team,</p>
<p>I hope this email finds you well.</p>
<p>I am writing from zenifix, a premium K-Beauty skincare brand based in Seoul. We would like to politely request your team's review of zenifix products for a potential retail partnership in your market.</p>
<p>We offer 14 core SKUs across two highly effective collections—our Noni Line (7 SKUs) and Ginkgo Line (7 SKUs). What truly sets zenifix apart is our exceptional ingredient concentration. Our formulations feature natural Noni and Ginkgo extracts <strong>ranging from 21% to 58% (210,000 ppm – 580,000 ppm)</strong> depending on the SKU. We differentiate our products through this uncompromising raw material content rather than generic marketing claims.</p>"""
            }
        },
        "마케팅 에이전시 (협업)": {
            "태국어": {
                "subject": "[ข้อเสนอความร่วมมือ] สกินแคร์ K-Beauty พรีเมียมจาก zenifix",
                "body": "<p>เรียน ทีมงานการตลาด,</p>\n<p>เราคือ zenifix แบรนด์สกินแคร์ระดับพรีเมียมจากโซล ประเทศเกาหลีใต้...</p>"
            }
        }
    }
    
# 탭(Tab)으로 수집 화면과 발송 화면 분리
tab1, tab2, tab3 = st.tabs(["📥 1. 이메일 수집 (Gathering)", "📧 2. 자동 발송 (Sending)", "📊 3. 데이터 대시보드 (통계)"])


# ==========================================
# [탭 1] 글로벌 이메일 수집 (기억 장치 적용 및 국가별 개별 저장)
# ==========================================
with tab1:
    st.header("글로벌 이메일 자동 수집기")
    
    # 💡 수집된 엑셀 데이터를 보관할 '기억 장치' 초기화
    if 'gathered_files' not in st.session_state:
        st.session_state.gathered_files = {}

    # 번역 지원 국가 사전
    COUNTRY_LANG_MAP = {
        "USA": "en", "UK": "en", "Australia": "en", "Canada": "en",
        "Ireland": "en", "New Zealand": "en", "India": "en", "Philippines": "en",
        "Germany": "de", "Austria": "de", "France": "fr", 
        "Japan": "ja", "Vietnam": "vi", "Thailand": "th", 
        "Spain": "es", "Mexico": "es", "UAE": "ar", "Italy": "it"
    }
    
    col1, col2 = st.columns(2)
    with col1:
        serp_api_key = st.text_input("SerpApi Key (필수)", type="password")
        countries_input = st.text_input("타깃 국가 (쉼표로 구분하여 복수 입력)", value="USA, Canada, Australia")
        selected_countries = [c.strip() for c in countries_input.split(",") if c.strip()]
        
    with col2:
        search_keyword = st.text_input("검색 키워드 (영문+현지어 듀얼 검색됨)", value="korean cosmetics distributor contact")
        page_count = st.number_input("검색어당 페이지 수", min_value=1, max_value=10, value=2)

    if st.button("🔍 이메일 수집 시작", type="primary"):
        if not serp_api_key or not selected_countries:
            st.error("SerpApi Key와 타깃 국가를 최소 1개 이상 입력해 주세요!")
        else:
            # 새로운 수집을 시작하면 이전 기억 장치를 비워줍니다.
            st.session_state.gathered_files = {}
            
            for country in selected_countries:
                st.markdown(f"### 🌍 {country} 수집 현황")
                status_text = st.empty()
                country_results = []
                
                lang_code = COUNTRY_LANG_MAP.get(country, "en")
                
                # 현지어 자동 번역
                try:
                    if lang_code != "en":
                        from deep_translator import GoogleTranslator
                        translated_keyword = GoogleTranslator(source='auto', target=lang_code).translate(search_keyword)
                    else:
                        translated_keyword = search_keyword
                except:
                    translated_keyword = search_keyword 
                
                search_queries = [f"{search_keyword} {country}"]
                if translated_keyword != search_keyword:
                    search_queries.append(f"{translated_keyword} {country}")

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
                                        country_results.append({
                                            "업체명": company_name, "국가명": country, "웹사이트": website_url,
                                            "담당자(유추)": "Cosmetics Purchasing Team", "이메일": emails[0], "수집상태": "대기중"
                                        })
                                    time.sleep(1) 
                        except Exception as e:
                            status_text.warning(f"검색 중 일시적 오류 발생: {e}")
                
                # 💡 수집 완료된 데이터를 기억 장치(session_state)에 저장합니다!
                if country_results:
                    df = pd.DataFrame(country_results)
                    df = df.drop_duplicates(subset=['이메일'], keep='first')
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    
                    st.session_state.gathered_files[country] = {
                        "count": len(df),
                        "data": output.getvalue()
                    }
                    status_text.success(f"🎉 {country} 수집 및 저장 완료! (순수 이메일 {len(df)}건)")
                else:
                    if not api_limit_hit:
                        status_text.warning(f"⚠️ {country}에서 수집된 이메일이 없습니다.")
                
                st.divider()

    # ==========================================
    # 💡 기억 장치에 저장된 엑셀 파일 다운로드 버튼 노출 (새로고침 방어)
    # ==========================================
    if st.session_state.gathered_files:
        st.subheader("📥 수집 완료된 국가별 엑셀 다운로드")
        for country, file_info in st.session_state.gathered_files.items():
            st.download_button(
                label=f"📥 {country} 엑셀 파일 다운로드 ({file_info['count']}건)", 
                data=file_info['data'],
                file_name=f"Zenifix_Buyers_{country}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_btn_{country}"
            )


# ==========================================
# [탭 2] 글로벌 이메일 자동 발송기 (수신거부 연동 완료)
# ==========================================
with tab2:
    st.header("글로벌 이메일 자동 발송기")
    
    if db_connected:
        st.success("✅ 구글 스프레드시트(zenifix_DB) 연동 완료! 발송 이력 및 수신거부 목록이 자동 연동됩니다.")

    # 버튼 디자인 CSS
    st.markdown("""
    <style>
    button[kind="primary"] {
        background-color: #03C75A !important;
        border-color: #03C75A !important;
        color: white !important;
    }
    button[kind="primary"]:hover {
        background-color: #028a3f !important;
        border-color: #028a3f !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # 계정 정보 입력
    col1, col2 = st.columns(2)
    with col1:
        login_email = st.text_input("개인 로그인 이메일 (예: zeni@wellsfnd.com)")
        app_password = st.text_input("16자리 앱 비밀번호", type="password")
    with col2:
        sender_email = st.text_input("발송자 이메일 (From: 공통메일)", value="zenifix@wellsfnd.com")
    
    st.divider()

    # --- [새로 추가된 기능] 수신거부 자동 동기화 버튼 ---
    st.subheader("🛡️ 수신거부(Unsubscribe) 자동 동기화")
    if st.button("🔄 내 수신함에서 '수신거부' 메일 찾아 DB에 자동 저장하기"):
        if not login_email or not app_password:
            st.warning("로그인 이메일과 앱 비밀번호를 먼저 입력해 주세요.")
        else:
            with st.spinner("수신함을 스캔하여 'Unsubscribe' 답장을 찾고 있습니다..."):
                try:
                    # 메일 수신함(IMAP) 로그인
                    mail = imaplib.IMAP4_SSL("imap.gmail.com")
                    mail.login(login_email, app_password)
                    mail.select("inbox")
                    
                    # 'Unsubscribe' 단어가 포함된 메일 검색
                    status, messages = mail.search(None, 'BODY "Unsubscribe"')
                    email_ids = messages[0].split()
                    
                    new_unsubs = 0
                    for e_id in email_ids:
                        res, msg_data = mail.fetch(e_id, '(RFC822)')
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                from_header = msg.get("From")
                                # 보낸 사람 주소에서 정확한 이메일만 추출
                                email_match = re.search(r'<(.+?)>', str(from_header))
                                sender = email_match.group(1) if email_match else str(from_header)
                                
                                # 구글 시트에 없는 이메일이면 새롭게 추가!
                                if sender not in blacklist_emails and "mailer-daemon" not in sender.lower():
                                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    blacklist_sheet.append_row([sender, current_time])
                                    blacklist_emails.append(sender)
                                    new_unsubs += 1
                                    
                    mail.logout()
                    st.toast(f"✅ 동기화 완료! {new_unsubs}명의 새로운 수신거부 이메일이 DB에 저장되었습니다.")
                except Exception as e:
                    st.error(f"동기화 중 오류 발생: {e}")

    st.divider()
    
    # 타깃, 언어, 이미지, 슬라이더 설정
    st.subheader("🎯 템플릿 및 발송 옵션 설정")
    with st.expander("➕ 새로운 타깃 그룹 및 언어 템플릿 추가하기"):
        new_target = st.text_input("새로운 타깃 그룹 이름 (예: VIP 바이어)")
        new_lang = st.text_input("새로운 발송 언어 (예: Spanish)")
        if st.button("템플릿 목록에 추가"):
            if new_target and new_lang:
                if new_target not in st.session_state.email_templates:
                    st.session_state.email_templates[new_target] = {}
                st.session_state.email_templates[new_target][new_lang] = {
                    "subject": f"[{new_target}] Partnership with zenifix",
                    "body": f"<p>Dear {new_target} Team,</p>\n<p>내용을 입력하세요.</p>"
                }
                st.success(f"'{new_target}' - '{new_lang}' 추가 완료!")
                st.rerun()

    col3, col4 = st.columns(2)
    with col3:
        target_type = st.selectbox("1. 타깃 그룹을 선택하세요", list(st.session_state.email_templates.keys()))
    with col4:
        available_languages = list(st.session_state.email_templates[target_type].keys())
        selected_language = st.selectbox("2. 발송 언어를 선택하세요", available_languages)

    use_image = st.radio("3. 본문 이미지 포함 여부", ["이미지 포함 (추천)", "텍스트만 발송 (이미지 없이)"], horizontal=True)
    img_url = ""
    if use_image == "이미지 포함 (추천)":
        img_url = st.text_input("이미지 URL 주소를 입력하세요", value="https://zenifix.net/img/zenifix_BrandDeck_main.png")
    
    delay_seconds = st.slider("4. 메일 발송 간격 조절 (스팸 방지용 대기 시간)", min_value=10, max_value=300, value=180, step=10)

    st.divider()

    # 이메일 편집기
    st.subheader("📝 이메일 미리보기 및 직접 편집")
    base_subject = st.session_state.email_templates[target_type][selected_language]["subject"]
    base_content = st.session_state.email_templates[target_type][selected_language]["body"]
    
    image_content = f'\n<p>zenifix 브랜드를 참고하실 수 있도록 아래에 간단한 이미지를 첨부하였습니다:</p>\n<p><img src="{img_url}" alt="zenifix Brand Overview" style="max-width: 800px; width: 100%; height: auto;"></p>\n' if (use_image == "이미지 포함 (추천)" and img_url) else ""
    
    footer_content = f"""
<p>To explore our complete Brand Deck, including full product details, current global sales channels, and our active SNS presence, please visit our official website:<br>
👉 <strong>Official Brand Deck: <a href="https://zenifix.net">https://zenifix.net</a></strong></p>
<p>If your team finds our brand suitable for your market after the initial review, please reply to this email. We would be happy to discuss further possibilities and details.</p>
<p>Thank you for your time and consideration.</p>
<p>Best regards,<br>
Global Partnership Team<br>
zenifix<br>
{sender_email}</p>
<p><small><i>*If you do not wish to receive further emails, please reply with 'Unsubscribe'.</i></small></p>
"""
    initial_html = base_content + image_content + footer_content
    
    edited_subject = st.text_input("📝 이메일 제목 (발송 전 자유롭게 수정 가능)", value=base_subject)
    edited_html_body = st.text_area("🔧 이메일 본문 (HTML 태그 및 텍스트 자유 수정)", value=initial_html, height=300)

    st.markdown("##### 👁️ 실제 수신자가 받아볼 이메일 미리보기")
    with st.container(border=True):
        st.markdown(f"**제목:** {edited_subject}")
        st.divider()
        components.html(edited_html_body, height=400, scrolling=True)

    st.divider()

    # 발송 컨트롤
    st.subheader("📅 발송 스케줄링 (예약 설정)")
    
    # 달력과 파일 업로드 창을 나란히 배치
    col_date, col_file = st.columns(2)
    with col_date:
        scheduled_date = st.date_input("발송을 원하는 예약 날짜를 선택하세요", min_value=datetime.today().date())
    with col_file:
        uploaded_file = st.file_uploader("수집 탭에서 다운로드한 '엑셀 파일'을 올려주세요.", type=["xlsx"])
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("🧪 내 메일로 테스트 1건 발송해 보기", use_container_width=True):
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
        # 🚀 버튼이 예약 등록 버튼으로 바뀌었습니다!
        if st.button("📅 지정한 날짜로 예약 발송 등록하기", type="primary", use_container_width=True):
            if not uploaded_file:
                st.error("엑셀 파일을 먼저 올려주세요!")
            elif not db_connected:
                st.error("구글 DB와 연결되지 않아 예약을 등록할 수 없습니다.")
            else:
                df = pd.read_excel(uploaded_file)
                st.info(f"총 {len(df)}명의 대상을 {scheduled_date} 발송 큐(Queue)에 등록합니다...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                success_count = 0
                skip_count = 0 # 수신거부 스킵 카운트
                
                for index, row in df.iterrows():
                    buyer_email = str(row.get('이메일', '')).strip()
                    buyer_country = str(row.get('국가명', '미확인'))
                    buyer_website = str(row.get('웹사이트', '미확인'))
                    
                    # --- [중요] DB 수신거부 필터링 (예약 단계에서 미리 차단!) ---
                    if buyer_email in blacklist_emails:
                        status_text.text(f"🚫 수신거부 대상 제외됨: {buyer_email}")
                        skip_count += 1
                        progress_bar.progress((index + 1) / len(df))
                        continue
                    
                    # 구글 시트에 넣을 준비
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    final_html = f"<html><body>{edited_html_body}</body></html>"
                    
                    try:
                        # 발송(smtplib) 대신 구글 시트 '발송예약' 탭에 정보 저장하기
                        queue_sheet.append_row([
                            str(scheduled_date), buyer_email, buyer_country, buyer_website, 
                            edited_subject, final_html, "대기중", target_type, current_time
                        ])
                        success_count += 1
                        status_text.text(f"✅ 예약 등록 완료 ({buyer_country}): {buyer_email}")
                    except Exception as e:
                        status_text.text(f"❌ DB 기록 실패: {e}")
                    
                    progress_bar.progress((index + 1) / len(df))
                    time.sleep(0.5) # API 과부하를 막기 위해 아주 짧게 휴식
                        
                st.success(f"🎉 총 {success_count}건 예약 완료! (수신거부 제외: {skip_count}명)\n지정하신 날짜({scheduled_date})에 시스템이 자동으로 발송합니다.")


# ==========================================
# [탭 3] 글로벌 발송 통계 대시보드
# ==========================================
with tab3:
    import plotly.express as px
    
    st.header("📊 글로벌 발송 데이터 대시보드")
    
    if not db_connected:
        st.warning("🚨 구글 스프레드시트와 연결되어 있지 않아 데이터를 불러올 수 없습니다.")
    else:
        # 데이터 새로고침 버튼
        if st.button("🔄 최신 데이터 불러오기", use_container_width=True):
            st.rerun()
            
        st.divider()
        
        try:
            # 구글 시트에서 모든 데이터 가져오기
            raw_data = db_sheet.get_all_values()
            
            # 헤더(첫 줄)를 제외하고 데이터가 1줄이라도 있는지 확인
            if len(raw_data) > 1:
                # 데이터를 표(DataFrame) 형태로 변환
                df_stats = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                
                # --- 1. 핵심 성과 지표 (KPI) 요약 ---
                total_sent = len(df_stats)
                total_unsubs = len(blacklist_emails) if 'blacklist_emails' in locals() else 0
                total_countries = df_stats['국가명'].nunique() if '국가명' in df_stats.columns else 0
                
                col1, col2, col3 = st.columns(3)
                col1.metric(label="🚀 총 발송 성공", value=f"{total_sent} 건")
                col2.metric(label="🌍 도달 국가 수", value=f"{total_countries} 개국")
                col3.metric(label="🚫 수신 거부 (블랙리스트)", value=f"{total_unsubs} 건")
                
                st.divider()
                
                # --- 2. 시각화 차트 ---
                col_chart1, col_chart2 = st.columns(2)
                
                with col_chart1:
                    st.subheader("📍 국가별 발송 비중")
                    if '국가명' in df_stats.columns:
                        country_counts = df_stats['국가명'].value_counts().reset_index()
                        country_counts.columns = ['국가명', '발송건수']
                        # 원형 차트 (Pie Chart)
                        fig_pie = px.pie(country_counts, values='발송건수', names='국가명', hole=0.4, 
                                         color_discrete_sequence=px.colors.sequential.Teal)
                        st.plotly_chart(fig_pie, use_container_width=True)
                        
                with col_chart2:
                    st.subheader("🎯 타깃 그룹별 발송 현황")
                    if '타깃유형' in df_stats.columns:
                        target_counts = df_stats['타깃유형'].value_counts().reset_index()
                        target_counts.columns = ['타깃유형', '발송건수']
                        # 막대 차트 (Bar Chart)
                        fig_bar = px.bar(target_counts, x='타깃유형', y='발송건수', text_auto=True,
                                         color='타깃유형', color_discrete_sequence=px.colors.qualitative.Pastel)
                        st.plotly_chart(fig_bar, use_container_width=True)

                st.divider()
                
                # --- 3. 최근 발송 이력 테이블 ---
                st.subheader("📝 최근 발송 이력 (최신 100건)")
                # 데이터를 거꾸로 뒤집어 최신순으로 정렬 후 100개만 노출
                st.dataframe(df_stats.iloc[::-1].head(100), use_container_width=True)
                
            else:
                st.info("💡 아직 구글 시트에 기록된 발송 데이터가 없습니다. 첫 콜드 메일을 발송하시면 통계가 자동으로 생성됩니다.")
                
        except Exception as e:
            st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
