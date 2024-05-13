"""General utility functions module.

This module contains a collection of general utility functions that are employed
across various projects to streamline and enhance coding practices. These functions
are designed to be reusable, efficient, and compatible with the common requirements
of different projects within the organization.
"""

import os
import time
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
import yaml
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font


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

    if (
        smtp_api_key is None
        and (smtp_username is None
        and smtp_password is None)
    ):
        raise ValueError(
            "Please provide "
            "either an smtp_api_key or "
            "(smtp_username and smtp_password)")

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

        server = smtplib.SMTP_SSL(smtp_host, smtp_port)
        server.ehlo()

        if smtp_api_key:
            server.login('apikey', smtp_api_key)
        else:
            server.login(smtp_username, smtp_password)

        server.sendmail(mail_from, mail_to, msg.as_string())
        server.close()
        print(f"Mail Has Been Sent To {mail_to} with Subject: {subject}")


def touch_excel(
        df: pd.DataFrame,
        file_path: str | Path,
        sheet_name: str = "Sheet1",
        add_df: pd.DataFrame = None,
):
    """this function updates or creates a new sheet with the given dataframe

    if the file does not exist yet, it will be created

    It can also concatenate two different DataFrames and merge them into one
    them write them in the given sheet

    Args:
        df (pd.DataFrame): 
            dataframe that is to be written to the sheet

        file_path (string): 
            path to the file to be written

        sheet_name (string, optional): 
            name of the sheet that is to updated. Defaults to "Sheet1".

        add_df (pd.DataFrame, optional): 
            other dataframe that is to be merged. Defaults to None.

    Raises:
        PermissionError: 
            When the file is opened by user and is currently being used. 
            Please Close the file
    """
    if isinstance(file_path, Path):
        file_path = str(file_path)

    if add_df is not None:
        df = pd.concat([df, add_df], ignore_index=True)

    try:
        if not os.path.exists(file_path):
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(
                file_path,
                mode="a",
                engine='openpyxl',
                if_sheet_exists='replace',
            ) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                time.sleep(5)
                # writer.close()
    except PermissionError as e:
        e.message = "File might be open. Close it."
        raise e


def get_config(
    file_name: str,
    is_global: bool = False,
    file_type: str = "json",
    encoding: str = "utf-8",
) -> dict:
    """
    Load configuration from a specified JSON or YAML file.

    This function reads the configuration from a file (either JSON or YAML)
    located either in the current working directory or in a global 
    "Project Configurations" directory in the user's home folder, depending 
    on the `is_global` flag.

    Args:
        file_name (str): The name of the file from which to load configurations.
        is_global (bool, optional): Flag to determine the location from which
            to load the configurations. If True, the configuration is loaded
            from the user's home directory. Otherwise, it's loaded from the
            current working directory. Defaults to False.
        file_type (str, optional): The type of the file to read from. Can be
            either 'json' or 'yaml'. Defaults to 'json'.
        encoding (str, optional): The encoding to use when opening the file.
            Defaults to 'utf-8'.

    Returns:
        dict: A dictionary containing the loaded configurations.

    Raises:
        FileNotFoundError: If the specified file does not exist in the expected
            location.
        ValueError: If the file_type is neither 'json' nor 'yaml'.
    """
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

    with open(config_path, 'r', encoding=encoding) as f:
        if file_type == "json":
            config = json.load(f)
        elif file_type == "yaml":
            config = yaml.load(f.read(), Loader=yaml.FullLoader)
        else:
            raise ValueError(f"File type {file_type} not supported")

    return config


def style_excel(path:str, sheet_name: str,  header_color = '3ce81e'):
    """
    Styles an Excel file by applying a header color and adjusting column widths.

    Parameters:
        path (str): The path to the Excel file.
        header_color (str, optional): The color of the header. Defaults to '3ce81e'.

    Returns:
        None

    Raises:
        AttributeError: If a column contains empty cells or is not formatted as expected.
    """
    input_workbook = load_workbook(path)
    input_worksheet = input_workbook[sheet_name]

    for column in input_worksheet.iter_cols():
        default_length = 10
        column_letter = column[0].column_letter
        column[0].fill = PatternFill(
            fill_type='solid', 
            start_color=header_color, 
            end_color=header_color,
        )
        column[0].font = Font(bold=True, size=13) 

        try:
            if len(str(column[0].value)) > default_length:
                default_length = len(column[0].value)

        except AttributeError as e:
            print(f'Column contains empty cells or the following column is not formatted as expected: {e}')

        adjusted_width = (default_length +2) * 1.2
        input_worksheet.column_dimensions[column_letter].width = adjusted_width

    input_workbook.save(path)
