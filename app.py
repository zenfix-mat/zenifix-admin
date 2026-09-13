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

# --- 새롭게 추가된 구글 시트(DB) 라이브러리 ---
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
try:
    # 스트림릿 Secrets에서 TOML 형식으로 저장된 열쇠를 불러옵니다.
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    gc = gspread.authorize(credentials)
    
    # 생성하신 구글 시트의 이름 'zenifix_DB'를 열어 첫 번째 시트를 준비합니다.
    db_sheet = gc.open("zenifix_DB").sheet1
    db_connected = True
except Exception as e:
    st.sidebar.error(f"구글 DB 연결 실패: {e}")
    st.sidebar.warning("발송 이력이 저장되지 않을 수 있습니다.")


# 탭(Tab)으로 수집 화면과 발송 화면 분리
tab1, tab2 = st.tabs(["📥 1. 이메일 수집 (Gathering)", "📧 2. 콜드 메일 자동 발송 (Sending)"])

# ==========================================
# [탭 1] 글로벌 이메일 수집 (SerpApi)
# ==========================================
with tab1:
    st.header("글로벌 이메일 자동 수집기")
    
    col1, col2 = st.columns(2)
    with col1:
        serp_api_key = st.text_input("SerpApi Key (필수)", type="password")
        target_country = st.text_input("타깃 국가 (예: USA, UK, Germany)", value="USA")
    with col2:
        search_keyword = st.text_input("검색 키워드", value="korean cosmetics distributor contact")
        page_count = st.number_input("검색할 페이지 수 (1페이지당 10개 결과)", min_value=1, max_value=10, value=2)

    if st.button("🔍 이메일 수집 시작", type="primary"):
        if not serp_api_key:
            st.error("SerpApi Key를 입력해 주세요!")
        else:
            with st.spinner("구글 검색 및 이메일 수집 중입니다... (약 1~2분 소요)"):
                results_data = []
                search_query_full = f"{search_keyword} {target_country}"
                
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

                # 검색 루프
                for page in range(page_count):
                    offset = page * 10
                    api_url = f"https://serpapi.com/search.json?engine=google&q={search_query_full}&start={offset}&api_key={serp_api_key}"
                    try:
                        response = requests.get(api_url).json()
                        if 'organic_results' in response:
                            for item in response['organic_results']:
                                company_name = item.get('title', '이름 없음')
                                website_url = item.get('link', '')
                                
                                if website_url.endswith('.pdf'): continue
                                
                                emails = extract_emails(website_url)
                                if emails:
                                    results_data.append({
                                        "업체명": company_name,
                                        "국가명": target_country,
                                        "웹사이트": website_url,
                                        "담당자(유추)": "Cosmetics Purchasing Team",
                                        "이메일": emails[0],
                                        "수집상태": "대기중"
                                    })
                                time.sleep(1) # 차단 방지
                    except Exception as e:
                        st.error(f"검색 중 오류 발생: {e}")

                # 수집 결과 화면 출력
                if results_data:
                    df = pd.DataFrame(results_data)
                    st.success(f"🎉 총 {len(df)}건의 이메일 수집 완료!")
                    st.dataframe(df)
                    
                    # 엑셀 다운로드 버튼
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📊 엑셀 파일 다운로드",
                        data=excel_data,
                        file_name=f"Zenifix_Buyers_{target_country}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("수집된 데이터가 없습니다. 키워드를 변경해 보세요.")


# ==========================================
# [탭 2] 글로벌 이메일 자동 발송기 (수신거부 자동화 탑재)
# ==========================================
with tab2:
    st.header("글로벌 이메일 자동 발송기")
    
    # --- [추가됨] 수신거부(Blacklist) DB 시트 준비 ---
    blacklist_emails = []
    if db_connected:
        try:
            # '수신거부'라는 이름의 탭(시트)이 있는지 찾고 데이터 가져오기
            blacklist_sheet = gc.open("zenifix_DB").worksheet("수신거부")
            blacklist_emails = blacklist_sheet.col_values(1) # 1열(A열) 이메일 목록
            st.success("✅ 구글 스프레드시트(zenifix_DB) 연동 완료! 발송 이력 및 수신거부 목록이 자동 연동됩니다.")
        except:
            # 탭이 없으면 자동으로 '수신거부' 탭을 새로 생성
            blacklist_sheet = gc.open("zenifix_DB").add_worksheet(title="수신거부", rows="1000", cols="2")
            blacklist_sheet.update_cell(1, 1, "이메일")
            blacklist_sheet.update_cell(1, 2, "등록일시")
            st.success("✅ 구글 시트에 [수신거부] 탭이 새로 생성되었습니다.")

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
                                email_match = re.search(r'<(.+?)>', from_header)
                                sender = email_match.group(1) if email_match else from_header
                                
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
        if st.button("🚀 전체 엑셀 리스트 대량 발송 시작", type="primary", use_container_width=True):
            if not uploaded_file or not login_email or not app_password:
                st.error("엑셀 파일, 로그인 이메일, 앱 비밀번호를 모두 입력해 주세요!")
            else:
                df = pd.read_excel(uploaded_file)
                st.info(f"총 {len(df)}명의 대상에게 발송을 시작합니다...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(login_email, app_password)
                    
                    success_count = 0
                    skip_count = 0 # 수신거부 스킵 카운트
                    
                    for index, row in df.iterrows():
                        buyer_email = row.get('이메일', '').strip()
                        buyer_country = row.get('국가명', '미확인')
                        buyer_website = row.get('웹사이트', '미확인')
                        
                        # --- [중요] DB 중복 및 수신거부 필터링 ---
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
                                    print(f"DB 기록 실패: {e}")
                                    
                            status_text.text(f"✅ 발송 완료 ({buyer_country}): {buyer_email}")
                        except:
                            status_text.text(f"❌ 발송 실패: {buyer_email}")
                        
                        progress_bar.progress((index + 1) / len(df))
                        
                        if index < len(df) - 1:
                            status_text.text(f"⏳ 스팸 방지를 위해 {delay_seconds}초 대기 중...")
                            time.sleep(delay_seconds)
                            
                    server.quit()
                    st.success(f"🎉 총 {success_count}건 발송 완료! (수신거부로 제외된 이메일: {skip_count}명)")
                    
                except Exception as e:
                    st.error(f"🚨 이메일 로그인 실패. 오류: {e}")
