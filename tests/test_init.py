import util
import pandas as pd
import os


def test_email_send_outlook():
    recipients = ["vinay.sagar@imarc.in"]
    subject = "Pytest Test Plain Email"
    message = "This is a plain test email."
    util.send_mail(subject, message, recipients, mode=util.EmailMode.OUTLOOK)


def test_email_send_api():
    util.smtp_host = os.environ["SMTP_HOST"]
    util.smtp_port = os.environ["SMTP_PORT"]
    util.smtp_username = os.environ["SMTP_USERNAME"]
    util.smtp_password = os.environ["SMTP_PASSWORD"]
    recipients = ["vinay.sagar@imarc.in"]
    subject = "Pytest Test Plain Email"
    message = "This is a plain test email."
    util.send_mail(subject, message, recipients, mode=util.EmailMode.API)


def test_email_type_html():
    recipients = ["vinay.sagar@imarc.in"]
    subject = "Pytest Test HTML Email"
    message = "<h1>HTML test email.</h1><p>This is a test email</p>"
    util.send_mail(subject, message, recipients, "html", mode=util.EmailMode.OUTLOOK)


def test_style_excel_invalid_sheet():
    error_occured = None
    df = pd.DataFrame({
        "this": ["a"],
        "is": ["b"],
        "an": ["c"],
        "invalid": ["d"]
    })

    os.makedirs("download_test", exist_ok=True)
    df.to_excel("download_test/workbook.xlsx", sheet_name="valid_name")

    try:
        util.style_excel("download_test/workbook.xlsx", sheet_name="invalid_name")
    except ValueError as e:
        error_occured = e
        print(e)
        assert "Could not find sheet" in str(error_occured)
    except Exception as e:
        error_occured = e
        print(e)
        assert False, "Unexpected error occurred."
    finally:
        os.remove("download_test/workbook.xlsx")


def test_style_excel_valid_sheet():
    error_occured = None
    df = pd.DataFrame({
        "this": ["a"],
        "is": ["b"],
        "an": ["c"],
        "invalid": ["d"]
    })

    os.makedirs("download_test", exist_ok=True)
    df.to_excel("download_test/workbook.xlsx", sheet_name="valid_name")

    try:
        util.style_excel("download_test/workbook.xlsx", sheet_name="valid_name")
        util.style_excel("download_test/workbook.xlsx", sheet_name=["valid_name"])
        util.style_excel("download_test/workbook.xlsx")
        assert True, "No error occurred."
    except ValueError:
        assert False, "Error occurred"
    finally:
        os.remove("download_test/workbook.xlsx")
