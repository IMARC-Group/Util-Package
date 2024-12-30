import util
import pandas as pd
import os
import pytest


def test_email_send_outlook():
    recipients = ["vinay.sagar@imarc.in", "forms@imarc.in"]
    # recipients = ["vinay.sagar@imarc.in"]
    subject = "Pytest Test Plain Email"
    message = "This is a plain test email."
    cc = ["vinay.sagar@imarc.in", "forms@imarc.in"]
    bcc = ["vinay.sagar@imarc.in", "forms@imarc.in"]
    util.send_mail(
        subject,
        message,
        recipients,
        mode=util.EmailMode.OUTLOOK,
        cc=cc,
        bcc=bcc,
    )


# def test_email_send_api():
#     util.smtp_host = os.environ["SMTP_HOST"]
#     util.smtp_port = os.environ["SMTP_PORT"]
#     util.smtp_username = os.environ["SMTP_USERNAME"]
#     util.smtp_password = os.environ["SMTP_PASSWORD"]
#     recipients = ["vinay.sagar@imarc.in", "forms@imarc.in"]
#     subject = "Pytest Test Plain Email"
#     message = "This is a plain test email."
#     util.send_mail(subject, message, recipients, mode=util.EmailMode.API)


def test_email_type_html():
    # recipients = ["vinay.sagar@imarc.in"]
    recipients = ["vinay.sagar@imarc.in", "forms@imarc.in"]
    subject = "Pytest Test HTML Email"
    message = "<h1>HTML test email.</h1><p>This is a test email</p>"
    cc = ["vinay.sagar@imarc.in", "forms@imarc.in"]
    bcc = ["vinay.sagar@imarc.in", "forms@imarc.in"]
    util.send_mail(
        subject,
        message,
        recipients,
        mode=util.EmailMode.OUTLOOK,
        mail_type="html",
        cc=cc,
        bcc=bcc,
    )


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

def test_fill_template_key_error():
    with pytest.raises(KeyError):
        util.fill_template(
            "tests/sample_template.md",
            verbose=True,
            data={},
        )

def test_fill_template_html():
    output = util.fill_template(
        "tests/sample_template.md",
        verbose=True,
        output_format="html",
        data={"sub_heading": "Hello, World!"}
    )
    assert "<h2>Hello, World!</h2>" == output


def test_fill_template_optional_data():
    output = util.fill_template(
        "tests/no_variable_tempalte.md",
        verbose=True,
        output_format="html",
    )
    assert "<h2>sub_heading</h2>" == output


def test_fill_template_plain():
    output = util.fill_template(
        "tests/sample_template.md",
        verbose=True,
        output_format="plain",
        data={"sub_heading": "Hello, World!"}
    )
    assert "## Hello, World!" == output
