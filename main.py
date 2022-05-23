import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
import openpyxl
import exchangelib

from dataclasses import dataclass


@dataclass
class Email:
    adress: str
    subject: str
    message: str


def main():
    import warnings
    warnings.simplefilter("ignore")

    filename = "data/20220512_Noten.xlsx"
    templates_folder = "templates"
    subject = "test"
    email_column = "Email"
    file = openpyxl.load_workbook(filename, data_only=True)

    with open("replacements.json", "r") as j:
        replacements = json.load(j)

    for i, sheet in enumerate(file.sheetnames):
        print(f"[{i + 1}] {sheet.title()}")

    try:
        choice = int(input("Choose a sheet: ")) - 1
    except ValueError:
        print("Invalid input")
        exit()

    if choice < 0 or choice > len(file.sheetnames):
        print("Invalid input")
        exit()

    ws = file[file.sheetnames[choice]]
    data = ws.values
    columns = next(data)[0:]
    df = pd.DataFrame(data, columns=columns)

    templates = [f for f in os.listdir(templates_folder)]

    for file in templates:
        print(f"[{templates.index(file) + 1}] {file}")

    try:
        choice = int(input("Choose a Template: ")) - 1
    except ValueError:
        print("Invalid input")
        exit()

    if choice < 0 or choice > len(templates):
        print("Invalid input")
        exit()

    with open(f"{templates_folder}/{templates[choice]}", "r") as f:
        template = f.read()

    emails = []
    for i, row in df.iterrows():
        message = template
        for column in columns:
            part = str(row[column])
            if part in replacements:
                part = replacements[part]
            message = message.replace(f"<{column}>", part)
        emails.append(Email(row[email_column], subject, message))

    print(emails)

    with open("data/credentials.json", "r") as j:
        credentials = json.load(j)

    sender_email = credentials["email"]
    password = credentials["password"]

    for email in emails:
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = email.adress
        message['Subject'] = email.subject
        message.attach(MIMEText(email.message, 'plain'))

        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(sender_email, password)
        text = message.as_string()
        s.sendmail(sender_email, email.adress, text)
        s.quit()

        print('Mail Sent')


if __name__ == "__main__":
    main()
