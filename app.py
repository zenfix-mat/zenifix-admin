import streamlit as st
import pandas as pd
import requests
import re
import time
import smtplib
from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO

# 웹페이지 기본 설정
st.set_page_config(page_title="zenifix Global Admin", page_icon="🚀", layout="wide")
st.title("🚀 zenifix Global B2B Admin Dashboard")

# 탭(Tab)으로 수집 화면과 발송 화면 분리
tab1, tab2 = st.tabs(["📥 1. 바이어 이메일 수집 (Gathering)", "📧 2. 콜드 메일 자동 발송 (Sending)"])

# ==========================================
# [탭 1] 바이어 이메일 수집 (SerpApi)
# ==========================================
with tab1:
    st.header("글로벌 바이어 이메일 자동 수집기")
    
    col1, col2 = st.columns(2)
    with col1:
        serp_api_key = st.text_input("SerpApi Key (필수)", type="password")
        target_country = st.text_input("타깃 국가 (예: USA, UK, Germany)", value="USA")
    with col2:
        search_keyword = st.text_input("검색 키워드", value="korean cosmetics distributor contact")
        page_count = st.number_input("검색할 페이지 수 (1페이지당 10개 결과)", min_value=1, max_value=10, value=2)

    if st.button("🔍 바이어 수집 시작", type="primary"):
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
# [탭 2] 콜드 메일 자동 발송 (SMTP)
# ==========================================
with tab2:
    st.header("글로벌 바이어 이메일 자동 발송기")
    
    col3, col4 = st.columns(2)
    with col3:
        login_email = st.text_input("개인 로그인 이메일 (예: zeni@wellsfnd.com)")
        app_password = st.text_input("16자리 앱 비밀번호", type="password")
        
    with col4:
        sender_email = st.text_input("발송자 이메일 (From: 공통메일)", value="zenifix@wellsfnd.com")
        img_url = st.text_input("브랜드 덱 요약 이미지 URL (필수입력)", value="https://zenifix.net/img/zenifix_BrandDeck_main.png")
        
    uploaded_file = st.file_uploader("수집 탭에서 다운로드한 '엑셀 파일'을 올려주세요.", type=["xlsx"])
    
    if st.button("🚀 이메일 발송 시작", type="primary"):
        if not uploaded_file or not login_email or not app_password:
            st.error("엑셀 파일, 로그인 이메일, 앱 비밀번호를 모두 입력해 주세요!")
        else:
            df = pd.read_excel(uploaded_file)
            st.info(f"총 {len(df)}명의 바이어에게 발송을 시작합니다...")
            
            # 진행 상태 바
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(login_email, app_password)
                
                success_count = 0
                for index, row in df.iterrows():
                    buyer_email = row['이메일']
                    
                    msg = MIMEMultipart()
                    msg['From'] = f"zenifix Team <{sender_email}>"
                    msg['To'] = buyer_email
                    msg['Subject'] = "[Partnership Proposal] Premium K-Beauty: 580,000ppm High-Concentration Skincare by zenifix"
                    msg.add_header('reply-to', sender_email)
                    
                    html_body = f"""
                    <html>
                        <body>
                            <p>Dear Cosmetics Purchasing Team,</p>
                            <p>I hope this email finds you well.</p>
                            <p>I am writing from zenifix, a premium K-Beauty skincare brand based in Seoul. We would like to politely request your team's review of zenifix products for a potential retail partnership in your market.</p>
                            <p>We offer 14 core SKUs across two highly effective collections—our Noni Line (7 SKUs) and Ginkgo Line (7 SKUs). What truly sets zenifix apart is our exceptional ingredient concentration. Our formulations feature natural Noni and Ginkgo extracts <strong>ranging from 21% to 58% (210,000 ppm – 580,000 ppm)</strong> depending on the SKU. We differentiate our products through this uncompromising raw material content rather than generic marketing claims.</p>
                            <p>We have included a brief image below for your reference regarding the zenifix brand:</p>
                            <p><img src="{img_url}" alt="zenifix Brand Overview" style="max-width: 800px; width: 100%; height: auto;"></p>
                            <p>To explore our complete Brand Deck, including full product details, current global sales channels, and our active SNS presence, please visit our official website:<br>
                            👉 <strong>Official Brand Deck: <a href="https://zenifix.net">https://zenifix.net</a></strong></p>
                            <p>If your team finds our brand suitable for your market after the initial review, please reply to this email. We would be happy to discuss further possibilities and details.</p>
                            <p>Thank you for your time and consideration.</p>
                            <p>Best regards,<br>
                            Global Partnership Team<br>
                            zenifix<br>
                            {sender_email}</p>
                            <p><small><i>*If you do not wish to receive further emails, please reply with 'Unsubscribe'.</i></small></p>
                        </body>
                    </html>
                    """
                    msg.attach(MIMEText(html_body, 'html'))
                    
                    try:
                        server.send_message(msg)
                        success_count += 1
                        status_text.text(f"✅ 발송 성공: {buyer_email}")
                    except:
                        status_text.text(f"❌ 발송 실패: {buyer_email}")
                    
                    # 프로그레스 바 업데이트
                    progress_bar.progress((index + 1) / len(df))
                    
                    # 마지막 메일이 아니면 3분 대기 (진짜 시스템처럼 180초 대기)
                    if index < len(df) - 1:
                        status_text.text("⏳ 스팸 방지를 위해 3분(180초) 대기 중...")
                        time.sleep(180)
                        
                server.quit()
                st.success(f"🎉 총 {success_count}건의 실전 콜드 메일 발송이 안전하게 완료되었습니다!")
                
            except Exception as e:
                st.error(f"🚨 이메일 로그인 실패. 앱 비밀번호를 다시 확인해 주세요. 오류: {e}")
