import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import mailer


def test_clean_product_removes_quantity_prefix():
    assert mailer.clean_product("2 * Wireless Mouse") == "Wireless Mouse"
    assert mailer.clean_product("Wireless Mouse") == "Wireless Mouse"
    assert mailer.clean_product(None) == ""


def test_read_customers_skips_duplicates():
    path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "sample_customers.xlsx")
    customers = mailer.read_customers(path)
    order_ids = [c[0] for c in customers]
    assert len(customers) == 4
    assert order_ids.count("DEMO-1003") == 1


def test_log_prevents_resend(tmp_path):
    log = str(tmp_path / "log.csv")
    mailer.write_log("DEMO-1001", "a@example.com", "sent", log_file=log)
    mailer.write_log("DEMO-1002", "b@example.com", "failed: TimeoutError", log_file=log)
    done = mailer.already_sent(log_file=log)
    assert done == {"DEMO-1001"}


def test_message_has_product_and_store_name():
    msg = mailer.build_message("me@example.com", "c@example.com", "Subject", "Phone Case")
    body = msg.get_content()
    assert "Phone Case" in body
    assert mailer.STORE_NAME in body
