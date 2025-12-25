import re
import pandas as pd

# ========== CẤU HÌNH ==========
LOG_FILE = "mail_error_graph_app.log"   # đổi thành tên file txt/log của bạn
OUTPUT_EXCEL = "emails_retry.xlsx"      # file excel xuất ra
EMAIL_COLUMN_NAME = "Email"

# Regex bắt email
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_emails_from_log(log_path):
    emails = []

    # cố gắng đọc utf-8, nếu lỗi có thể đổi sang encoding khác
    with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            # Tìm tất cả email trong dòng
            found = EMAIL_REGEX.findall(line)
            for e in found:
                e_clean = e.strip()
                if e_clean:
                    emails.append(e_clean)

    # Loại trùng
    unique_emails = sorted(set(emails))
    return unique_emails

def save_emails_to_excel(emails, excel_path, col_name="Email"):
    if not emails:
        print("Không tìm thấy email nào để xuất.")
        return

    df = pd.DataFrame({col_name: emails})
    df.to_excel(excel_path, index=False)
    print(f"Đã lưu {len(emails)} email vào file: {excel_path}")

if __name__ == "__main__":
    emails = extract_emails_from_log(LOG_FILE)
    print(f"Tìm được {len(emails)} email lỗi:")
    for e in emails:
        print(" -", e)

    save_emails_to_excel(emails, OUTPUT_EXCEL, EMAIL_COLUMN_NAME)
