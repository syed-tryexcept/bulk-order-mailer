"""Bulk order-update mailer.

Reads customers from an Excel file and sends each one a short, templated
email over SMTP. Built for e-commerce order notices (for example a delivery
timeline for overseas products).

Safety features:
  * Preview mode sends nothing.
  * Test mode sends only to your own address.
  * Live mode needs you to type YES.
  * Every send is logged, and orders already sent are skipped on re-runs.
"""
import csv
import os
import re
import smtplib
import ssl
import time
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr

from dotenv import load_dotenv
from openpyxl import load_workbook

load_dotenv()

# ---- Settings (override any of them in your .env file) ---------------------
EXCEL_FILE = os.getenv("EXCEL_FILE", "sample_customers.xlsx")
LOG_FILE = os.getenv("LOG_FILE", "sent_log.csv")
STORE_NAME = os.getenv("STORE_NAME", "Demo Store")
SUBJECT = os.getenv("EMAIL_SUBJECT", f"Order Delivery Timeline \u2013 {STORE_NAME}")
DELIVERY_TIME = os.getenv("DELIVERY_TIME", "30-35 days")
PRODUCT_NAME_OVERRIDE = os.getenv("PRODUCT_NAME_OVERRIDE", "")  # empty = use the Excel item name
DELAY_SECONDS = float(os.getenv("DELAY_SECONDS", "3"))

COL_ORDER = "Customer order id"
COL_EMAIL = "Customer email"
COL_ITEM = "Item Name"


def get_smtp_settings():
    """Read SMTP credentials from the environment. Fails with a clear message."""
    keys = ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASSWORD"]
    missing = [k for k in keys if not os.getenv(k)]
    if missing:
        raise SystemExit(
            "Missing settings in .env: " + ", ".join(missing)
            + "\nCopy .env.example to .env and fill in your SMTP details."
        )
    return {
        "host": os.getenv("SMTP_HOST"),
        "port": int(os.getenv("SMTP_PORT")),
        "user": os.getenv("SMTP_USER"),
        "password": os.getenv("SMTP_PASSWORD"),
    }


def make_body(product):
    return f"""Dear Customer,

Greetings from {STORE_NAME}!

Your {product} is an overseas product and will take approximately {DELIVERY_TIME} to be delivered.

Thank you for your patience and for choosing {STORE_NAME}.

Best regards,
{STORE_NAME} Team
"""


def clean_product(item):
    """'2 * Wireless Mouse' -> 'Wireless Mouse'."""
    item = str(item or "").strip()
    return re.sub(r"^\d+\s*\*\s*", "", item)


def read_customers(path=None):
    """Return a list of unique (order_id, email, product) tuples."""
    wb = load_workbook(path or EXCEL_FILE, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() for h in next(rows)]
    i_order = header.index(COL_ORDER)
    i_email = header.index(COL_EMAIL)
    i_item = header.index(COL_ITEM)

    customers = []
    seen = set()
    for row in rows:
        order_id = str(row[i_order] or "").strip()
        email = str(row[i_email] or "").strip()
        product = PRODUCT_NAME_OVERRIDE or clean_product(row[i_item])
        if not order_id or not email:
            continue
        if not product:
            print("WARNING: no product name, skipped:", order_id)
            continue
        key = (order_id, email.lower())
        if key in seen:
            continue
        seen.add(key)
        customers.append((order_id, email, product))
    return customers


def already_sent(log_file=None):
    """Order ids that already have a 'sent' row in the log."""
    log_file = log_file or LOG_FILE
    done = set()
    if os.path.exists(log_file):
        with open(log_file, newline="", encoding="utf-8") as f:
            for r in csv.reader(f):
                if len(r) >= 4 and r[3] == "sent":
                    done.add(r[1])
    return done


def write_log(order_id, email, status, log_file=None):
    log_file = log_file or LOG_FILE
    new_file = not os.path.exists(log_file)
    with open(log_file, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["time", "order_id", "email", "status"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), order_id, email, status])


def build_message(smtp_user, to_email, subject, product, bcc=None):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr((f"{STORE_NAME} Support Team", smtp_user))
    msg["To"] = to_email
    if bcc:
        msg["Bcc"] = bcc
    msg.set_content(make_body(product))
    return msg


def send(msg, smtp):
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp["host"], smtp["port"], context=context, timeout=30) as server:
        server.login(smtp["user"], smtp["password"])
        server.send_message(msg)


def run_preview(customers):
    if customers:
        print("\n--- Sample email (first customer) ---")
        print("Subject:", SUBJECT)
        print(make_body(customers[0][2]))
    print("--- Recipients ---")
    done = already_sent()
    for n, (order_id, email, product) in enumerate(customers, 1):
        mark = "(already sent)" if order_id in done else ""
        print(n, order_id, email, "|", product, mark)


def run_test(customers):
    smtp = get_smtp_settings()
    me = input("Your own email address: ").strip()
    for order_id, _email, product in customers[:3]:
        send(build_message(smtp["user"], me, "[TEST] " + SUBJECT, product), smtp)
        print("Test sent for order", order_id, "->", me)
        time.sleep(DELAY_SECONDS)


def run_live(customers):
    smtp = get_smtp_settings()
    done = already_sent()
    todo = [c for c in customers if c[0] not in done]
    print("Already sent:", len(customers) - len(todo), "| To send now:", len(todo))
    if input(f"Type YES to send to {len(todo)} customers: ").strip() != "YES":
        print("Cancelled.")
        return
    for order_id, email, product in todo:
        try:
            send(build_message(smtp["user"], email, SUBJECT, product, bcc=smtp["user"]), smtp)
            write_log(order_id, email, "sent")
            print("SENT  ", order_id, email, "|", product)
        except Exception as e:
            write_log(order_id, email, "failed: " + type(e).__name__)
            print("FAILED", order_id, email, type(e).__name__, e)
        time.sleep(DELAY_SECONDS)
    print(f"Done. Details: {LOG_FILE}")


def main():
    customers = read_customers()
    print("Unique customers found:", len(customers))
    print()
    print("1 = Preview (sends nothing)")
    print("2 = Test (first 3 emails go to your own address)")
    print("3 = Live (sends to customers)")
    choice = input("Option (1/2/3): ").strip()

    if choice == "1":
        run_preview(customers)
    elif choice == "2":
        run_test(customers)
    elif choice == "3":
        run_live(customers)
    else:
        print("Invalid option.")


if __name__ == "__main__":
    main()
