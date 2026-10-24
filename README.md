# Bulk Order Mailer

A small Python tool that sends templated order-update emails (for example a
delivery timeline for overseas products) to customers listed in an Excel file.

Built from a real e-commerce operations need: contacting many customers about
the same notice, without sending the same email twice.

## Features
- **Three modes:** Preview (sends nothing), Test (sends to your own address), Live (needs you to type `YES`)
- **No double sends:** every send is logged, and orders already sent are skipped on re-runs
- **Duplicate rows removed:** the same order and email is only emailed once
- **Safe secrets:** SMTP credentials live in a `.env` file that is never committed
- **Rate limit:** a configurable delay between emails
- **Tests:** `pytest` tests for parsing, logging and message building

## Setup
```bash
pip install -r requirements.txt
cp .env.example .env      # Windows: copy .env.example .env
```
Edit `.env` and fill in your SMTP host, port, user and password.

## Input file
An Excel file with these columns (see `sample_customers.xlsx`):

| Customer order id | Customer email | Item Name |
|---|---|---|
| DEMO-1001 | customer1@example.com | 1 * Wireless Mouse |

A leading quantity like `2 * ` in the item name is removed automatically.
The sample file uses fake `example.com` addresses.

## Run
```bash
python mailer.py
```
Choose `1` (Preview) first, then `2` (Test), and only then `3` (Live).

## Test
```bash
pytest
```

## Configuration
Optional settings in `.env`: `STORE_NAME`, `DELIVERY_TIME`, `EXCEL_FILE`,
`DELAY_SECONDS`, `EMAIL_SUBJECT`, `PRODUCT_NAME_OVERRIDE`.

## Privacy
Do not commit `.env`, real customer spreadsheets or `sent_log.csv`.
They are already listed in `.gitignore`.
