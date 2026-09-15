import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import pytz
import time

# 1. 깃허브 금고(Secrets)에서 열쇠 꺼내오기
gcp_secret_json = os.environ.get("GCP_SERVICE_ACCOUNT")
login_email = os.environ.get("MAIL_ID")
app_password = os.environ.get("MAIL_PW")
sender_email = "zenifix@wellsfnd.com" 

# 2. 구글 시트 연결
creds_dict = json.loads(gcp_secret_json)
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
gc = gspread.authorize(credentials)

db_sheet = gc.open("zenifix_DB").sheet1
queue_sheet = gc.open("zenifix_DB").worksheet("발송예약")

# 💡 [핵심 패치] 기준 시간을 한국(Seoul) 시간으로 고정!
kst = pytz.timezone('Asia/Seoul')
now_str = datetime.now(kst).strftime("%Y-%m-%d %H:%M")
print(f"[{now_str}] 🤖 시간 지정 예약 발송 로봇 가동 시작...")

# 3. 지정된 시간이 지난 '대기중' 큐(Queue) 찾기
records = queue_sheet.get_all_records()
pending_rows = []
for idx, row in enumerate(records, start=2):
    status = str(row.get("상태", ""))
    reserve_time = str(row.get("예약일", "")) 
    
    # 💡 [핵심 패치] 상태가 '대기중'이고, 예약 시간이 현재 시간과 같거나 지났을 경우에만 색출
    if status == "대기중" and reserve_time <= now_str:
        pending_rows.append((idx, row))

if not pending_rows:
    print("✅ 현재 시간이 된 발송 예약 건이 없습니다. 로봇을 종료합니다.")
    exit()

print(f"🚀 발송 대상: 총 {len(pending_rows)}건")

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
        current_time = datetime.now(kst).strftime("%Y-%m-%d %H:%M:%S")

        queue_sheet.update_cell(row_num, 7, "발송완료")
        db_sheet.append_row([current_time, buyer_email, country, website, target_type, "English", "성공"])
        print(f"✅ 발송 성공: {buyer_email}")
        
    except Exception as e:
        print(f"❌ 발송 실패: {buyer_email} ({e})")
        queue_sheet.update_cell(row_num, 7, "실패")

    # 스팸 방지 휴식 (60초 단위 분산 발송)
    time.sleep(60)

server.quit()
print("🎉 예약 발송 작업이 모두 완료되었습니다!")
