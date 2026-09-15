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

kst = pytz.timezone('Asia/Seoul')
now_str = datetime.now(kst).strftime("%Y-%m-%d %H:%M")
print(f"[{now_str}] 🤖 시간 지정 예약 발송 로봇 가동 시작...")

# 3. 지정된 시간이 지난 '대기중' 큐(Queue) 찾기
records = queue_sheet.get_all_records()
pending_rows = []
for idx, row in enumerate(records, start=2):
    status = str(row.get("상태", ""))
    reserve_time = str(row.get("예약일", "")) 
    
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

rows_to_delete = [] # 💡 지워야 할 엑셀 줄 번호를 담아둘 바구니

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

        # 💡 [핵심 패치] 메인 시트에 기록하고, 발송이 끝난 줄 번호를 바구니에 담습니다.
        db_sheet.append_row([current_time, buyer_email, country, website, target_type, "English", "성공"])
        rows_to_delete.append(row_num)
        print(f"✅ 발송 성공: {buyer_email}")
        
    except Exception as e:
        print(f"❌ 발송 실패: {buyer_email} ({e})")
        queue_sheet.update_cell(row_num, 7, "실패")

    time.sleep(60)

server.quit()

# 💡 [핵심 패치] 성공한 내역을 구글 시트에서 완전히 삭제합니다!
# (행이 꼬이지 않도록 맨 밑의 줄부터 거꾸로 지워 올라갑니다.)
for r in sorted(rows_to_delete, reverse=True):
    queue_sheet.delete_rows(r)

print("🎉 예약 발송 및 대기열 청소 작업이 모두 완료되었습니다!")
