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

    # 5. [새로 추가] 이메일 수집 예약(Queue) 탭 세팅
    try:
        gather_queue_sheet = gc.open("zenifix_DB").worksheet("수집예약")
    except:
        gather_queue_sheet = gc.open("zenifix_DB").add_worksheet(title="수집예약", rows="1000", cols="7")
        gather_queue_sheet.append_row(["예약일시", "타깃국가", "검색키워드", "페이지수", "상태", "수집건수", "등록일시"])
        
    # 6. [새로 추가] 수집된 바이어 결과 저장 탭 세팅
    try:
        gather_result_sheet = gc.open("zenifix_DB").worksheet("수집결과")
    except:
        gather_result_sheet = gc.open("zenifix_DB").add_worksheet(title="수집결과", rows="1000", cols="6")
        gather_result_sheet.append_row(["수집일시", "국가명", "업체명", "웹사이트", "이메일", "검색키워드"])

except Exception as e:
    st.sidebar.error(f"구글 DB 연결 실패: {e}")
    st.sidebar.warning("발송 이력 및 템플릿 저장이 작동하지 않을 수 있습니다.")


# ==========================================
# [영구 템플릿] 구글 시트에서 템플릿 불러오기
# ==========================================
if 'email_templates' not in st.session_state:
    templates = {}
    
    # 1. DB에 저장된 템플릿이 있으면 시트에서 우선적으로 모두 불러오기 (메인 로직)
    if db_connected and len(template_records) > 1:
        for row in template_records[1:]:
            if len(row) >= 4:
                tgt, lng, sub, bdy = row[0], row[1], row[2], row[3]
                if tgt not in templates:
                    templates[tgt] = {}
                templates[tgt][lng] = {"subject": sub, "body": bdy}
    else:
        # 2. 구글 시트가 완전히 비어있을 때 앱 오류를 막기 위한 '최소한의 기본 틀'
        # (기존의 길고 복잡한 HTML 하드코딩 문구는 모두 삭제했습니다.)
        templates = {
            "바이어 (유통/입점)": {
                "English": {
                    "subject": "[Partnership Proposal] Premium K-Beauty by zenifix",
                    "body": "<p>Dear Cosmetics Purchasing Team,</p>\n<p>내용을 입력해 주세요.</p>"
                }
            }
        }
        # 빈 시트에 최소 기본 틀 최초 기록
        if db_connected:
            template_sheet.append_row(["바이어 (유통/입점)", "English", templates["바이어 (유통/입점)"]["English"]["subject"], templates["바이어 (유통/입점)"]["English"]["body"]])

    st.session_state.email_templates = templates

# 탭 분리
tab1, tab2, tab3 = st.tabs(["📥 1. 이메일 수집 (Gathering)", "📧 2. 자동 발송 (Sending)", "📊 3. 데이터 대시보드 (통계)"])

# ==========================================
# [탭 1] 글로벌 이메일 수집 (스케줄링 예약 및 대기열 관리자)
# ==========================================
with tab1:
    st.header("글로벌 이메일 자동 수집기")
    
    COUNTRY_LANG_MAP = {
        "USA": "en", "UK": "en", "Australia": "en", "Canada": "en",
        "Germany": "de", "Austria": "de", "France": "fr", 
        "Japan": "ja", "Vietnam": "vi", "Thailand": "th", 
        "Spain": "es", "Mexico": "es", "UAE": "ar", "Italy": "it",
        "China": "zh-CN", "Taiwan": "zh-TW", "Russia": "ru",
        "Brazil": "pt", "Indonesia": "id", "Poland": "pl"
    }
    
    col1, col2 = st.columns(2)
    with col1:
        serp_api_key = st.text_input("SerpApi Key (필수)", type="password")
        countries_input = st.text_input("타깃 국가 (쉼표로 구분하여 복수 입력)", value="USA, UK, Germany")
        selected_countries = [c.strip() for c in countries_input.split(",") if c.strip()]
        
    with col2:
        search_keyword = st.text_input("검색 키워드 (영문+현지어 듀얼 검색됨)", value="korean cosmetics distributor contact")
        page_count = st.number_input("검색어당 페이지 수", min_value=1, max_value=10, value=2)

    st.divider()

    # --- 🕒 수집 스케줄링 (예약 설정) ---
    st.subheader("🕒 수집 예약 스케줄링")
    st.info("원하는 날짜와 시간을 지정해 두면, 클라우드 로봇이 알아서 구글을 검색하고 바이어 리스트를 모아옵니다.")
    
    col_date, col_time = st.columns(2)
    with col_date:
        gather_date = st.date_input("수집을 시작할 날짜", min_value=datetime.today().date())
    with col_time:
        gather_time = st.time_input("수집을 시작할 시간")
        
    scheduled_datetime = datetime.combine(gather_date, gather_time).strftime("%Y-%m-%d %H:%M")
    
    if st.button("📅 지정한 날짜/시간에 수집 예약하기", type="primary", use_container_width=True):
        if not serp_api_key or not selected_countries:
            st.error("SerpApi Key와 타깃 국가를 입력해 주세요!")
        elif not db_connected:
            st.error("구글 시트가 연결되지 않아 예약을 등록할 수 없습니다.")
        else:
            with st.spinner("예약 대기열에 등록 중입니다..."):
                success_count = 0
                import pytz
                kst = pytz.timezone('Asia/Seoul')
                current_time = datetime.now(kst).strftime("%Y-%m-%d %H:%M:%S")
                
                try:
                    gather_queue_sheet = gc.open("zenifix_DB").worksheet("수집예약")
                    for country in selected_countries:
                        gather_queue_sheet.append_row([
                            scheduled_datetime, country, search_keyword, page_count, 
                            "대기중", "0", current_time
                        ])
                        success_count += 1
                    st.success(f"🎉 총 {success_count}개 지역의 수집 예약이 완료되었습니다!")
                except Exception as e:
                    st.error(f"DB 기록 실패: {e}")

    # ==========================================
    # --- 📊 수집 예약 대기열 및 완료 현황 관리 ---
    # ==========================================
    st.divider()
    st.subheader("📋 수집 예약 대기열 및 현황 관리")
    
    if db_connected:
        try:
            gather_queue_sheet = gc.open("zenifix_DB").worksheet("수집예약")
            g_queue_records = gather_queue_sheet.get_all_records()
            
            if g_queue_records:
                df_g_queue = pd.DataFrame(g_queue_records)
                # 구글 시트의 실제 행 번호(2번째 줄부터 시작)를 매핑해 줍니다.
                df_g_queue['시트행번호'] = df_g_queue.index + 2
                
                # '대기중' 또는 '진행중'인 항목
                df_pending = df_g_queue[df_g_queue['상태'].isin(['대기중', '진행중'])].copy()
                # '수집완료' 항목
                df_done = df_g_queue[df_g_queue['상태'] == '수집완료'].copy()
                
                # -----------------------------------
                # 1. 수집 대기 중 (관리 & 수동 실행)
                # -----------------------------------
                st.markdown(f"**⏳ 수집 대기 및 진행 중 ({len(df_pending)}건)**")
                if not df_pending.empty:
                    df_pending.insert(0, '삭제선택', False)
                    # 데이터 에디터로 체크박스 UI 제공
                    edited_pending = st.data_editor(
                        df_pending[['삭제선택', '시트행번호', '예약일시', '타깃국가', '검색키워드', '상태']],
                        hide_index=True,
                        use_container_width=True,
                        disabled=['시트행번호', '예약일시', '타깃국가', '검색키워드', '상태'],
                        column_config={"삭제선택": st.column_config.CheckboxColumn("삭제", default=False)}
                    )
                    
                    col_m1, col_m2 = st.columns(2)
                    with col_m1:
                        # 🗑️ [기능 1] 선택 삭제
                        if st.button("🗑️ 선택한 예약 삭제하기"):
                            rows_to_delete = edited_pending[edited_pending['삭제선택'] == True]['시트행번호'].tolist()
                            if rows_to_delete:
                                # 시트 행이 꼬이지 않도록 맨 아랫줄부터 거꾸로 지웁니다.
                                for r in sorted(rows_to_delete, reverse=True):
                                    gather_queue_sheet.delete_rows(r)
                                st.success("선택한 예약이 시트에서 삭제되었습니다.")
                                st.rerun()
                            else:
                                st.warning("삭제할 항목을 먼저 체크해 주세요.")
                    
                    with col_m2:
                        # 🚀 [기능 2] 수동 즉시 수집
                        if st.button("🚀 시간이 지난 예약 '수동으로 강제 수집'"):
                            import pytz
                            kst = pytz.timezone('Asia/Seoul')
                            now_kst = datetime.now(kst).strftime("%Y-%m-%d %H:%M")
                            
                            # 시간이 지났고 대기중인 것만 색출
                            past_due = df_pending[(df_pending['상태'] == '대기중') & (df_pending['예약일시'] <= now_kst)]
                            
                            if past_due.empty:
                                st.info("현재 시간이 지나 대기 중인 항목이 없습니다.")
                            else:
                                gather_result_sheet = gc.open("zenifix_DB").worksheet("수집결과")
                                
                                for idx, row in past_due.iterrows():
                                    r_num = row['시트행번호']
                                    t_loc = row['타깃국가']
                                    t_kw = row['검색키워드']
                                    p_cnt = int(row.get('페이지수', 2))
                                    
                                    # [상태 업데이트: 대기중 -> 진행중]
                                    gather_queue_sheet.update_cell(r_num, 5, "진행중")
                                    st.toast(f"[{t_loc}] 수집 진행중...")
                                    
                                    # 언어 변환 및 검색 (gatherer.py 로직)
                                    lang_code = "en"
                                    for c, c_code in COUNTRY_LANG_MAP.items():
                                        if c.lower() in t_loc.lower():
                                            lang_code = c_code
                                            break
                                    try:
                                        if lang_code != "en":
                                            from deep_translator import GoogleTranslator
                                            trans_kw = GoogleTranslator(source='auto', target=lang_code).translate(t_kw)
                                        else:
                                            trans_kw = t_kw
                                    except:
                                        trans_kw = t_kw
                                        
                                    queries = [f"{t_kw} {t_loc}"]
                                    if trans_kw != t_kw: queries.append(f"{trans_kw} {t_loc}")
                                    
                                    collected_cnt = 0
                                    api_limit = False
                                    
                                    def extract_emails(url):
                                        try:
                                            resp = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
                                            soup = BeautifulSoup(resp.text, 'html.parser')
                                            ep = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
                                            emls = re.findall(ep, soup.get_text(separator=' '))
                                            for link in soup.find_all('a'):
                                                href = link.get('href')
                                                if href and href.startswith('mailto:'):
                                                    emls.append(href.replace('mailto:', '').split('?')[0])
                                            return list({e.lower() for e in emls if not e.endswith(('.png', '.jpg', '.gif'))})
                                        except: return []

                                    for q in queries:
                                        if api_limit: break
                                        for p in range(p_cnt):
                                            url = f"https://serpapi.com/search.json?engine=google&q={q}&start={p*10}&api_key={serp_api_key}"
                                            try:
                                                res = requests.get(url).json()
                                                if 'error' in res:
                                                    api_limit = True
                                                    break
                                                if 'organic_results' in res:
                                                    for item in res['organic_results']:
                                                        link = item.get('link', '')
                                                        if link.endswith('.pdf'): continue
                                                        found = extract_emails(link)
                                                        if found:
                                                            c_time = datetime.now(kst).strftime("%Y-%m-%d %H:%M:%S")
                                                            gather_result_sheet.append_row([c_time, t_loc, item.get('title',''), link, found[0], t_kw])
                                                            collected_cnt += 1
                                                    time.sleep(1)
                                            except: pass
                                            
                                    # [상태 업데이트: 진행중 -> 수집완료]
                                    gather_queue_sheet.update_cell(r_num, 5, "수집완료")
                                    gather_queue_sheet.update_cell(r_num, 6, str(collected_cnt))
                                
                                st.success("🎉 강제 수동 수집이 모두 완료되었습니다!")
                                st.rerun()
                else:
                    st.info("대기 중인 수집 예약이 없습니다.")
                        
                # -----------------------------------
                # 2. 수집 완료 내역 (자동 분류 & 비우기 기능)
                # -----------------------------------
                st.markdown("<br>", unsafe_allow_html=True)
                col_d1, col_d2 = st.columns([3, 1])
                with col_d1:
                    st.markdown(f"**✅ 수집 완료 내역 ({len(df_done)}건)**")
                with col_d2:
                    # 🧹 [기능 3] 완료 내역 시트에서 자동 삭제
                    if st.button("🧹 완료 목록 삭제"):
                        if not df_done.empty:
                            rows_to_clear = df_done['시트행번호'].tolist()
                            for r in sorted(rows_to_clear, reverse=True):
                                gather_queue_sheet.delete_rows(r)
                            st.toast("완료된 내역이 모두 정리되었습니다.")
                            st.rerun()

                if not df_done.empty:
                    st.dataframe(df_done[['예약일시', '타깃국가', '수집건수', '상태']].iloc[::-1], hide_index=True, use_container_width=True)
                else:
                    st.info("최근 완료된 수집 내역이 없습니다.")
                    
            else:
                st.markdown("아직 등록된 수집 예약 데이터가 없습니다.")
                
        except Exception as e:
            st.warning(f"대기열 정보를 불러오는 중 오류 발생: {e}")


# ==========================================
# [탭 2] 글로벌 콜드 메일 자동 발송 (Excel-Free & 다이렉트 DB 연동)
# ==========================================
with tab2:
    st.header("글로벌 바이어 이메일 자동 발송기")
    
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

    # 템플릿 영구 저장 버튼
    if st.button("💾 현재 수정한 제목과 본문을 '현재 템플릿'으로 영구 저장", type="primary", use_container_width=True):
        st.session_state.email_templates[target_type][selected_language]["subject"] = edited_subject
        st.session_state.email_templates[target_type][selected_language]["body"] = edited_html_body
        
        if db_connected:
            try:
                records = template_sheet.get_all_values()
                found_row_idx = -1
                for i, row in enumerate(records):
                    if i > 0 and row[0] == target_type and row[1] == selected_language:
                        found_row_idx = i + 1
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

    # --- 관리자 계정 설정 및 수신거부 관리 ---
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
    
    # =========================================================================
    # [새로운 핵심 기능] 엑셀-프리(Excel-Free) 다이렉트 DB 연동 및 휴먼 검수 
    # =========================================================================
    st.subheader("📥 [Excel-Free] 수집된 바이어 DB 직접 불러오기 및 검수")
    st.caption("로봇이 구글 시트(수집결과 탭)에 모아둔 바이어 목록을 바로 불러와서 확인하고 발송할 수 있습니다.")
    
    # 1. DB에서 데이터 불러오기 버튼
    if st.button("🔄 로봇이 수집한 최신 바이어 DB 불러오기"):
        if db_connected:
            try:
                gather_result_sheet = gc.open("zenifix_DB").worksheet("수집결과")
                result_records = gather_result_sheet.get_all_values() # 👈 get_all_records() 대신 get_all_values() 사용
                
                # 💡 [핵심 패치] 데이터가 '헤더(1줄)'를 제외하고 실제로 존재하는지 확인합니다.
                if result_records and len(result_records) > 1:
                    # 첫 줄(헤더)을 컬럼명으로 삼아 DataFrame 생성
                    st.session_state.loaded_buyers_df = pd.DataFrame(result_records[1:], columns=result_records[0])
                    st.success(f"🎉 총 {len(st.session_state.loaded_buyers_df)}명의 바이어 데이터를 성공적으로 불러왔습니다!")
                else:
                    # 💡 데이터가 비어있을 때 뜨는 친절하고 세련된 안내 메시지
                    st.info("📭 아직 로봇이 수집을 완료한 바이어 데이터가 없습니다. [탭 1]에서 먼저 수집 스케줄을 예약해 주세요.")
                    if 'loaded_buyers_df' in st.session_state:
                        del st.session_state.loaded_buyers_df
            except Exception as e:
                st.error(f"수집결과 데이터를 불러오는 중 일시적인 오류가 발생했습니다: {e}")
        else:
            st.error("구글 DB와 연결되어 있지 않습니다.")

    # 2. 데이터 에디터 (휴먼 검수 및 선택)
    if 'loaded_buyers_df' in st.session_state and not st.session_state.loaded_buyers_df.empty:
        df_buyers = st.session_state.loaded_buyers_df.copy()
        
        # '선택'이라는 체크박스 열을 맨 앞에 추가 (기본값 True)
        if '선택' not in df_buyers.columns:
             df_buyers.insert(0, '선택', True)
             
        st.markdown("##### 🧐 발송 전 바이어 휴먼 검수")
        st.caption("발송을 원하지 않는 이메일(예: 고객센터, 아마존 등)은 아래 표에서 **'선택' 체크박스를 해제**해 주세요.")
        
        # 사용자가 직접 표에서 체크박스를 수정할 수 있는 data_editor 사용
        edited_df = st.data_editor(
            df_buyers,
            hide_index=True,
            use_container_width=True,
            disabled=["수집일시", "국가명", "업체명", "웹사이트", "이메일", "검색키워드"], # 다른 열은 수정 불가
            column_config={
                "선택": st.column_config.CheckboxColumn("발송 선택", help="체크 해제 시 발송 대상에서 제외됩니다.", default=True)
            }
        )
        
        # 체크된 항목만 최종 발송 대상으로 필터링
        final_df = edited_df[edited_df['선택'] == True].copy()
        st.info(f"선택된 최종 발송 대상: **{len(final_df)}명** / 전체: {len(edited_df)}명")
        
        st.markdown("<br>", unsafe_allow_html=True)

        # --- 발송 스케줄링 및 엑셀 기반 예약 발송 (기존과 동일하지만, 대상이 final_df로 변경됨) ---
        col_date, col_delay = st.columns(2)
        with col_date:
            scheduled_date = st.date_input("📅 달력에서 예약 발송 일자를 선택하세요", min_value=datetime.today().date())
        with col_delay:
            delay_seconds = st.slider("메일 발송 간격 조절 (즉시 발송 시 적용, 단위: 초)", min_value=10, max_value=300, value=180, step=10)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
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
                if final_df.empty:
                    st.error("발송할 바이어를 선택해 주세요!")
                elif not login_email or not app_password:
                    st.error("로그인 이메일, 앱 비밀번호를 모두 입력해 주세요!")
                else:
                    st.info(f"총 {len(final_df)}명의 대상에게 즉시 발송을 시작합니다...")
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.starttls()
                        server.login(login_email, app_password)
                        
                        success_count = 0
                        skip_count = 0 
                        
                        for index, row in final_df.iterrows():
                            buyer_email = str(row.get('이메일', '')).strip()
                            buyer_country = str(row.get('국가명', '미확인'))
                            buyer_website = str(row.get('웹사이트', '미확인'))
                            
                            if buyer_email in blacklist_emails:
                                status_text.text(f"🚫 수신거부 대상 제외됨: {buyer_email}")
                                skip_count += 1
                                progress_bar.progress((index + 1) / len(final_df))
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
                            
                            progress_bar.progress((index + 1) / len(final_df))
                            if index < len(final_df) - 1:
                                status_text.text(f"⏳ 스팸 방지를 위해 {delay_seconds}초 대기 중...")
                                time.sleep(delay_seconds)
                                
                        server.quit()
                        st.success(f"🎉 총 {success_count}건 실시간 발송 완료! (수신거부 스킵: {skip_count}명)")
                        
                    except Exception as e:
                        st.error(f"🚨 이메일 로그인 실패. 오류: {e}")

        with col_btn3:
            if st.button("📅 지정한 날짜로 예약 등록", type="primary", use_container_width=True):
                if final_df.empty:
                    st.error("예약할 바이어를 선택해 주세요!")
                elif not db_connected:
                    st.error("구글 DB와 연결되지 않아 예약을 등록할 수 없습니다.")
                else:
                    st.info(f"총 {len(final_df)}명의 대상을 {scheduled_date} 예약 대기열에 등록합니다...")
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    success_count = 0
                    skip_count = 0 
                    
                    try:
                        existing_queue_emails = queue_sheet.col_values(2) 
                    except:
                        existing_queue_emails = []

                    for index, row in final_df.iterrows():
                        buyer_email = str(row.get('이메일', '')).strip()
                        buyer_country = str(row.get('국가명', '미확인'))
                        buyer_website = str(row.get('웹사이트', '미확인'))
                        
                        if buyer_email in blacklist_emails:
                            status_text.text(f"🚫 수신거부 대상 제외됨: {buyer_email}")
                            skip_count += 1
                            progress_bar.progress((index + 1) / len(final_df))
                            continue
                            
                        if buyer_email in existing_queue_emails:
                            status_text.text(f"⚠️ 이미 예약된 바이어 제외됨: {buyer_email}")
                            skip_count += 1
                            progress_bar.progress((index + 1) / len(final_df))
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
                        
                        progress_bar.progress((index + 1) / len(final_df))
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
