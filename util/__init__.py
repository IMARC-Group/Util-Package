import os
import time
import json
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


smtp_host: str | None = None
smtp_port: int | None = None
smtp_username: str | None = None
smtp_password: str | None = None
smtp_api_key: str | None = None


def send_mail(
        subject: str,
        message: str,
        recipients: list[str],
        mail_type: str = "plain"
):
    """sends mail

    Args:
        mail_type:
        recipients:
        subject (str): subject to send
        message (str): message to send
        recipients (string): email addresses to send to
    """
    global smtp_host
    global smtp_port
    global smtp_username
    global smtp_password
    global smtp_api_key

    if smtp_api_key is None and (smtp_username is None and smtp_password is None):
        raise ValueError("Please provide either an smtp_api_key or (smtp_username and smtp_password)")

    for name in recipients:
        time.sleep(1)
        mail_from = 'alan.baker@imarcgroup.info'
        mail_to = str(name)

        msg = MIMEMultipart()
        msg['From'] = mail_from
        msg['To'] = mail_to
        msg['Subject'] = str(subject)

        html_part = MIMEText(str(message), mail_type)
        msg.attach(html_part)

        server = smtplib.SMTP_SSL('smtp.sendgrid.net', 465)
        server.ehlo()

        if smtp_api_key:
            server.login('apikey', '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')
        else:
            server.login(smtp_username, smtp_password)

        server.sendmail(mail_from, mail_to, msg.as_string())
        server.close()
        print(f"Mail Has Been Sent To {mail_to} with Subject: {subject}")


def touch_excel(
        df: pd.DataFrame,
        file_path: str,
        sheet_name: str = "Sheet1",
        add_df: pd.DataFrame = None,
):
    """this function updates or creates a new sheet with the given dataframe

    if the file does not exist yet, it will be created

    It can also concatenate two different DataFrames and merge them into one
    them write them in the given sheet

    Args:
        df (pd.DataFrame): dataframe that is to be written to the sheet
        file_path (string): path to the file to be written
        sheet_name (string, optional): name of the sheet that is to updated. Defaults to "Sheet1".
        add_df (pd.DataFrame, optional): other dataframe that is to be merged. Defaults to None.

    Raises:
        Exception: When the file is opened by user and is currently being used. Please Close the file
    """

    if add_df is not None:
        df = pd.concat([df, add_df], ignore_index=True)

    try:
        if not os.path.exists(file_path):
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(file_path, mode="a", engine='openpyxl', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                time.sleep(5)
                # writer.close()
    except PermissionError:
        raise Exception("File might be open. Close it.")


def get_config(file_name: str, is_global: bool = False) -> dict:
    if is_global:
        base_path = os.path.join(os.path.expanduser('~'), "Project Configurations")
    else:
        base_path = os.getcwd()

    config_path = os.path.join(
        base_path,
        file_name,
    )

    if not os.path.exists(config_path):
        raise FileNotFoundError(f"{file_name} Not found at {config_path}")

    with open(config_path, 'r') as f:
        config = json.load(f)

    return config
