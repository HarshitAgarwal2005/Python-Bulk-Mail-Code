import os
import re
import time
import smtplib
import pandas as pd
from email.message import EmailMessage
from dotenv import load_dotenv
# pip install pandas openpyxl python-dotenv

EXCEL_FILE_PATH = r"D:\Bulk Mails\emails.xlsx"

FROM_EMAIL = "prep4btech@gmail.com"
TO_EMAIL = "harshitagarwal25806@gmail.com"

SUBJECT = "Your Email Subject Here"

BATCH_SIZE = 100
DELAY_SECONDS = 10

EMAIL_COLUMN = "email"
LOG_FILE = "sent_batches_log.txt"

# =========================
# HTML EMAIL CONTENT --- Replace your content in the HTML Format below
# =========================

HTML_CONTENT = """
<html>
  <body>
    <h2>Hello Student,</h2>

    <p>This is your email content written in HTML.</p>

    <p>
      You can add links, images, buttons, tables, and formatting here.
    </p>

    <p>Regards,<br>
    Prep4BTech Team</p>
  </body>
</html>
"""


# =========================
# EMAIL VALIDATION
# =========================

def is_valid_email(email):
    email_regex = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return re.match(email_regex, email) is not None


def load_and_validate_emails(file_path):
    df = pd.read_excel(file_path)

    df.columns = df.columns.str.strip().str.lower()

    if EMAIL_COLUMN not in df.columns:
        raise ValueError(f"Excel file must contain an '{EMAIL_COLUMN}' column.")

    raw_emails = (
        df[EMAIL_COLUMN]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
        .drop_duplicates()
        .tolist()
    )

    valid_emails = []
    rejected_emails = []

    for email in raw_emails:
        if is_valid_email(email):
            valid_emails.append(email)
        else:
            rejected_emails.append({
                "email": email,
                "reason": "Invalid email format"
            })

    return valid_emails, rejected_emails


# =========================
# LOGGING
# =========================

def write_log(message):
    with open(LOG_FILE, "a", encoding="utf-8") as file:
        file.write(message + "\n")


def log_rejected_emails(rejected_emails):
    if not rejected_emails:
        return

    write_log("\n========== REJECTED EMAILS ==========")

    for item in rejected_emails:
        email = item["email"]
        reason = item["reason"]

        print(f"Rejected: {email} | Reason: {reason}")
        write_log(f"REJECTED | Email: {email} | Reason: {reason}")


def log_sent_batch(batch_number, emails):
    write_log(f"\n========== BATCH {batch_number} SENT ==========")

    for email in emails:
        write_log(f"SENT | Batch: {batch_number} | Email: {email}")


def log_failed_batch(batch_number, emails, error):
    write_log(f"\n========== BATCH {batch_number} FAILED ==========")
    write_log(f"ERROR: {error}")

    for email in emails:
        write_log(f"FAILED | Batch: {batch_number} | Email: {email}")


# =========================
# EMAIL SENDING
# =========================

def send_email_batch(sender_email, app_password, to_email, bcc_emails, subject, html_content):
    msg = EmailMessage()

    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.set_content("This email requires an HTML-compatible email client.")
    msg.add_alternative(html_content, subtype="html")

    all_recipients = [to_email] + bcc_emails

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, app_password)

        rejected_by_server = smtp.send_message(
            msg,
            from_addr=sender_email,
            to_addrs=all_recipients
        )

    return rejected_by_server


# =========================
# MAIN PROGRAM
# =========================

def main():
    load_dotenv()

    gmail_app_password = os.getenv("GMAIL_APP_PASSWORD")

    if not gmail_app_password:
        raise ValueError("GMAIL_APP_PASSWORD not found. Please add it in your .env file.")

    open(LOG_FILE, "w", encoding="utf-8").close()

    print("Reading Excel file...")
    valid_emails, rejected_emails = load_and_validate_emails(EXCEL_FILE_PATH)

    print(f"Valid emails found: {len(valid_emails)}")
    print(f"Rejected emails found: {len(rejected_emails)}")

    write_log("========== BULK MAIL REPORT ==========")
    write_log(f"Valid emails found: {len(valid_emails)}")
    write_log(f"Rejected emails found: {len(rejected_emails)}")

    log_rejected_emails(rejected_emails)

    if not valid_emails:
        print("No valid emails available for sending.")
        write_log("\nNo valid emails available for sending.")
        return

    total_batches = (len(valid_emails) + BATCH_SIZE - 1) // BATCH_SIZE

    for batch_index in range(total_batches):
        batch_number = batch_index + 1

        start = batch_index * BATCH_SIZE
        end = start + BATCH_SIZE

        batch_emails = valid_emails[start:end]

        print("\n" + "=" * 50)
        print(f"Sending Batch {batch_number}/{total_batches}")
        print(f"Emails in this batch: {len(batch_emails)}")
        print("=" * 50)

        for email in batch_emails:
            print(f"Adding to BCC: {email}")

        try:
            rejected_by_server = send_email_batch(
                sender_email=FROM_EMAIL,
                app_password=gmail_app_password,
                to_email=TO_EMAIL,
                bcc_emails=batch_emails,
                subject=SUBJECT,
                html_content=HTML_CONTENT
            )

            if rejected_by_server:
                print(f"Batch {batch_number} sent, but some emails were rejected by server.")

                write_log(f"\n========== BATCH {batch_number} PARTIALLY SENT ==========")

                rejected_server_emails = set(rejected_by_server.keys())

                for email in batch_emails:
                    if email in rejected_server_emails:
                        reason = rejected_by_server[email]
                        print(f"Rejected by server: {email} | Reason: {reason}")
                        write_log(f"REJECTED BY SERVER | Batch: {batch_number} | Email: {email} | Reason: {reason}")
                    else:
                        print(f"Sent successfully: {email}")
                        write_log(f"SENT | Batch: {batch_number} | Email: {email}")

            else:
                print(f"Batch {batch_number} sent successfully.")

                for email in batch_emails:
                    print(f"Sent successfully: {email}")

                log_sent_batch(batch_number, batch_emails)

        except Exception as error:
            print(f"Batch {batch_number} failed.")
            print(f"Error: {error}")

            log_failed_batch(batch_number, batch_emails, error)

        if batch_index < total_batches - 1:
            print(f"\nWaiting {DELAY_SECONDS} seconds before next batch...")
            time.sleep(DELAY_SECONDS)

    print("\nBulk mailing completed.")
    print(f"Final report saved in: {LOG_FILE}")

    write_log("\n========== BULK MAIL COMPLETED ==========")


if __name__ == "__main__":
    main()