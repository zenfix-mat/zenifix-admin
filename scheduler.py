import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time

# 1. 깃허브 금고(Secrets)에서 열쇠 꺼내오기
gcp_secret_json = os.environ.get("GCP_SERVICE_ACCOUNT")
login_email = os.environ.get("MAIL_ID")
app_password = os.environ.get("MAIL_PW")
sender_email = "zenifix@wellsfnd.com" # 대표 공통 발신 메일

# 2. 구글 시트 연결
creds_dict = json.loads(gcp_secret_json)
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
gc = gspread.authorize(credentials)

db_sheet = gc.open("zenifix_DB").sheet1
queue_sheet = gc.open("zenifix_DB").worksheet("발송예약")

today_str = datetime.now().strftime("%Y-%m-%d")
print(f"[{today_str}] 🤖 예약 발송 로봇 가동 시작...")

# 3. 오늘 발송할 '대기중' 큐(Queue) 찾기
records = queue_sheet.get_all_records()
pending_rows = []
for idx, row in enumerate(records, start=2): # 엑셀은 2번째 줄부터 데이터 시작
    # 예약일이 오늘이거나 과거이고, 상태가 '대기중'인 것만 색출
    if str(row.get("상태")) == "대기중" and str(row.get("예약일")) <= today_str:
        pending_rows.append((idx, row))

if not pending_rows:
    print("✅ 오늘 발송할 예약 건이 없습니다. 로봇을 종료합니다.")
    exit()

print(f"🚀 오늘 발송 대상: 총 {len(pending_rows)}건")

# 4. 이메일 서버 연결 및 발송 시작
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(login_email, app_password)

for row_num, row_data in pending_rows:
    buyer_email = str(row_data.get("이메일"))
    subject = str(row_data.get("제목"))
    body_html = str(row_data.get("본문"))
    country = str(row_data.get("국가명"))
    website = str(row_data.get("웹사이트"))
    target_type = str(row_data.get("타깃유형"))

    msg = MIMEMultipart()
    msg['From'] = f"zenifix Team <{sender_email}>"
    msg['To'] = buyer_email
    msg['Subject'] = subject
    msg.add_header('reply-to', sender_email)
    msg.attach(MIMEText(body_html, 'html'))

    try:
        server.send_message(msg)
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 발송예약 시트 업데이트 (대기중 -> 발송완료)
        queue_sheet.update_cell(row_num, 7, "발송완료")
        
        # 메인 통계 시트(Sheet1)에 이력 꼼꼼히 기록
        db_sheet.append_row([current_time, buyer_email, country, website, target_type, "English", "성공"])
        print(f"✅ 발송 성공: {buyer_email}")
        
    except Exception as e:
        print(f"❌ 발송 실패: {buyer_email} ({e})")
        queue_sheet.update_cell(row_num, 7, "실패")

    # 스팸 차단 방지를 위해 메일 1통당 60초 휴식 (로봇이 알아서 기다립니다)
    time.sleep(60)

server.quit()
print("🎉 오늘의 예약 발송 작업이 모두 완료되었습니다!")
