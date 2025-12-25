import time
from datetime import datetime
import json
import requests
import msal
import pandas as pd
import re

EMAIL_REGEX = re.compile(r"^[\w\.-]+@[\w\.-]+\.\w+$")
INVALID_EMAIL_LOG = "invalid_email.log"
# ==========================
# CẤU HÌNH FILE EXCEL
# ==========================
EXCEL_FILE_PATH = "emails_retry.xlsx"   # đường dẫn file Excel
EMAIL_COLUMN_NAME = "Email"                  # tên cột chứa email

# Tên công ty hiển thị ở trường FROM
COMPANY_NAME = "Aigreeting Company"

# ==========================
# CẤU HÌNH ỨNG DỤNG GRAPH (APP-ONLY)
# ==========================
CLIENT_ID = "XXXX"   # Application (client) ID từ Azure
TENANT_ID = "XXXX"   # Directory (tenant) ID từ Azure
CLIENT_SECRET = "XXXX"      # ⚠ THAY BẰNG VALUE THẬT

# App-only dùng scope .default (lấy theo Application permissions đã gán)
SCOPES = ["https://graph.microsoft.com/.default"]

# Mailbox sẽ đứng tên gửi (UPN hoặc primary email)
SENDER_EMAIL = "admin@aigreetings.com.vn"

# ==========================
# CẤU HÌNH GỬI MAIL
# ==========================
BATCH_SIZE = 500
SLEEP_BETWEEN_BATCH = 60  # nghỉ 60 giây giữa mỗi batch

SUBJECT = "BÃI GFORTUNE THÔNG BÁO"

HTML_BODY = """
<html>
<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; background-color: #f4f4f4; padding: 20px;">
    <div style="max-width: 650px; margin: auto; background: #ffffff; padding: 30px; border: 1px solid #e0e0e0; border-radius: 8px;">
        <h3 style="color: #1a5da4; border-bottom: 2px solid #1a5da4; padding-bottom: 10px;">THÔNG BÁO</h3>
        
        <p style="white-space: pre-line;">
        Công ty CP GREATING FORTUNE CONTAINER VIỆT NAM xin thông báo, như thông lệ các năm để phục vụ cho việc hạch toán và quyết toán doanh thu – chi phí của năm yêu cầu quý khách hàng có container phát sinh việc hoàn lại tiền phí dịch vụ nâng hạ cont, thời điểm từ 20/12 kể về trước vui lòng liên hệ với bãi để làm thủ tục hoàn tiền. Hạn hoàn tiền đến hết ngày 29/12. Sau ngày trên chúng tôi sẽ tiến hành khóa sổ và không hoàn tiền với những container của khoảng thời gian đã thông báo như trên.
        
        Quý khách hàng vui lòng làm như hướng dẫn. 
        
        Trong quá trình thao tác nếu phát sinh vấn đề, vui lòng liên hệ với Phòng Kế toán tại Bãi để được hướng dẫn cụ thể.
        
        Nhân viên phụ trách Ms Ngọc, điện thoại 0906046646
        
        Trân trọng cảm ơn
        </p>
      
    </div>
</body>
</html>
"""

# ==========================
# LOGGING
# ==========================
ERROR_LOG_FILE = "mail_error_graph_app.log"
SUCCESS_LOG_FILE = "mail_success_graph_app.log"


def log_error(message: str):
    with open(ERROR_LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()} | ERROR | {message}\n")


def log_success(message: str):
    with open(SUCCESS_LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()} | SUCCESS | {message}\n")


# ==========================
# OAUTH2 – CLIENT CREDENTIALS (APP-ONLY)
# ==========================
def get_access_token_app():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET,
    )

    result = app.acquire_token_silent(SCOPES, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=SCOPES)

    if "access_token" not in result:
        raise RuntimeError(f"Không lấy được access token: {result}")

    return result["access_token"]


# ==========================
# TEST: APP CÓ NHÌN THẤY USER GỬI MAIL KHÔNG?
# ==========================
def test_sender_access(access_token: str):
    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    resp = requests.get(url, headers=headers)
    print(f"Test GET /users/{SENDER_EMAIL} => {resp.status_code}")
    if resp.status_code != 200:
        log_error(
            f"Không truy cập được user {SENDER_EMAIL}. "
            f"status={resp.status_code}, body={resp.text}"
        )
        print(f"Chi tiết lỗi: {resp.text}")
        raise RuntimeError(
            f"App không có quyền truy cập user {SENDER_EMAIL} hoặc user không tồn tại."
        )


# ==========================
# HÀM CHIA BATCH
# ==========================
def chunk_list(lst, size):
    for i in range(0, len(lst), size):
        yield lst[i:i + size]

# Truong hop loi gui tung mail mot
def send_single_graph_app(access_token: str, email: str):
    payload = {
        "message": {
            "subject": SUBJECT,
            "body": {"contentType": "HTML", "content": HTML_BODY},
            "toRecipients": [{"emailAddress": {"address": email}}],
        },
        "saveToSentItems": True
    }
    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    return requests.post(url, headers=headers, data=json.dumps(payload))

def _send_one_message(access_token: str, email_list):
    bcc_recipients = [{"emailAddress": {"address": e}} for e in email_list]

    payload = {
        "message": {
            "subject": SUBJECT,
            "body": {"contentType": "HTML", "content": HTML_BODY},
            "from": {"emailAddress": {"name": COMPANY_NAME, "address": SENDER_EMAIL}},
            "sender": {"emailAddress": {"name": COMPANY_NAME, "address": SENDER_EMAIL}},
            "bccRecipients": bcc_recipients
        },
        "saveToSentItems": True
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    # thêm timeout để tránh treo
    return requests.post(url, headers=headers, data=json.dumps(payload), timeout=60)

# ==========================
# GỬI 1 BATCH QUA GRAPH (APP-ONLY)
# ==========================
# Batch nào fail vì 1 email → tự động “chẻ đôi” để lôi ra đúng email gây lỗi
# Các email còn lại vẫn gửi được
# Email lỗi được ghi vào invalid_email.log
def send_batch_graph_app(access_token: str, recipients_batch, batch_index: int):
    email_list = [r.get("email", "") for r in recipients_batch]
    email_list = [re.sub(r"\s+", "", e) for e in email_list if e and e.strip()]
    email_list = [e for e in email_list if e]

    if not email_list:
        msg = f"[Batch {batch_index}] Không có email hợp lệ."
        print(msg, flush=True)
        log_error(msg)
        return

    resp = _send_one_message(access_token, email_list)

    if resp.status_code == 202:
        msg = f"[Batch {batch_index}] Gửi thành công {len(email_list)} khách."
        print(msg, flush=True)
        log_success(msg)
        return

    # Nếu lỗi invalid recipients, chia đôi để tìm email làm hỏng
    body = resp.text or ""
    if "ErrorInvalidRecipients" in body and len(email_list) > 1:
        mid = len(email_list) // 2
        left = [{"email": e} for e in email_list[:mid]]
        right = [{"email": e} for e in email_list[mid:]]
        log_error(f"[Batch {batch_index}] Batch fail, tách đôi để tìm email lỗi. status={resp.status_code} body={body}")
        send_batch_graph_app(access_token, left, f"{batch_index}.1")
        send_batch_graph_app(access_token, right, f"{batch_index}.2")
        return

    # Nếu chỉ còn 1 email mà vẫn lỗi => chính nó lỗi, log vào invalid file
    if len(email_list) == 1:
        bad = email_list[0]
        log_error(f"[Batch {batch_index}] Email bị Graph reject: {bad} | status={resp.status_code} | {body}")
        with open(INVALID_EMAIL_LOG, "a", encoding="utf-8") as f:
            f.write(f"{bad}\n")
        print(f"[Batch {batch_index}] Email lỗi: {bad}", flush=True)
        return

    # Lỗi khác
    err = f"[Batch {batch_index}] Lỗi gửi mail, status={resp.status_code}, body={body}"
    print(err, flush=True)
    log_error(err)

# ==========================
# ĐỌC DANH SÁCH EMAIL TỪ EXCEL
# ==========================
def load_recipients_from_excel(path: str, email_col: str = "Email"):
    try:
        df = pd.read_excel(path)
    except Exception as ex:
        log_error(f"Lỗi đọc file Excel: {ex}")
        raise

    if email_col not in df.columns:
        msg = f"Không tìm thấy cột '{email_col}' trong file Excel."
        log_error(msg)
        raise RuntimeError(msg)

    recipients = []
    invalid_emails = []

    for val in df[email_col].dropna():
        raw_email = str(val)
        # 🔥 Xóa mọi khoảng trắng (đầu, giữa, cuối)
        email = re.sub(r"\s+", "", raw_email)
        if not email:
            continue

        if not EMAIL_REGEX.match(email):
            invalid_emails.append(email)
            log_error(f"❌ Email sai định dạng, bỏ qua: {email}")
            with open(INVALID_EMAIL_LOG, "a", encoding="utf-8") as f:
                f.write(f"{email}\n")
            continue

        # Email hợp lệ → thêm vào batch
        recipients.append({"email": email})

    total_valid = len(recipients)
    total_invalid = len(invalid_emails)

    print(f"Đọc từ Excel: {total_valid} email hợp lệ, {total_invalid} email sai định dạng.")
    log_success(
        f"Đọc từ Excel '{path}': {total_valid} email hợp lệ, {total_invalid} email sai định dạng."
    )

    return recipients
# ==========================
# MAIN
# ==========================
def send_email_to_customers_via_graph_app():
    # 1. Đọc danh sách email từ Excel
    recipients = load_recipients_from_excel(EXCEL_FILE_PATH, EMAIL_COLUMN_NAME)
    
    print(f"Tổng khách cần gửi: {len(recipients)}")

    if not recipients:
        print("Không có email nào trong file Excel.")
        return

    # 2. Lấy access token
    try:
        access_token = get_access_token_app()
        print("Đã lấy access token (app-only).")
    except Exception as ex:
        log_error(f"Lỗi lấy access token: {ex}")
        print(f"Lỗi lấy access token: {ex}")
        return

    # 3. Test quyền mailbox gửi
    try:
        test_sender_access(access_token)
        print(f"App có quyền truy cập mailbox {SENDER_EMAIL}.")
    except Exception as ex:
        print(f"Lỗi khi test quyền truy cập user: {ex}")
        return

    # 4. Gửi theo batch
    for idx, batch in enumerate(chunk_list(recipients, BATCH_SIZE), start=1):
        print(f"--- Gửi batch {idx} ({len(batch)} khách) ---")
        send_batch_graph_app(access_token, batch, idx)

        if SLEEP_BETWEEN_BATCH > 0:
            print(f"Chờ {SLEEP_BETWEEN_BATCH} giây trước batch tiếp theo...")
            time.sleep(SLEEP_BETWEEN_BATCH)


# ==========================
# CHẠY
# ==========================
if __name__ == "__main__":
    send_email_to_customers_via_graph_app()
