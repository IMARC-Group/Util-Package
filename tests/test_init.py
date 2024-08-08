import util

def test_email_send_plain():
    recipients = ["vinay.sagar@imarc.in"]
    subject = "Pytest Test Plain Email"
    message = "This is a plain test email."
    util.send_mail(subject, message, recipients)


def test_email_send_html():
    recipients = ["vinay.sagar@imarc.in"]
    subject = "Pytest Test HTML Email"
    message = "<h1>HTML test email.</h1><p>This is a test email</p>"
    util.send_mail(subject, message, recipients, "html")
