#!/usr/bin/env python

import os
import csv

import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

EMAILS_FILE = "emails.csv" # the excel-pdf-generator writes to this file
FROM = "foo@example.com" # TODO
PASSWORD = "test123" # TODO
SMTP_SERVER = "smtp.gmail.com" # TODO
SMTP_PORT = 465

def main():
    with open(EMAILS_FILE) as file:
        reader = csv.reader(file)
        next(reader)  # skip header row
        send_emails(reader)

def send_emails(reader):
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
        server.login(FROM, PASSWORD)

        for company, receiver_email, pdf_path in reader:
            print(f"Sending email to {email}")

            subject = "Sąskaitos išrašas"
            body = ""

            text = prepare_email(receiver_email, subject, body, [pdf_path])
            server.sendmail(FROM, receiver_email, text)

def prepare_email(to, subject, body, attachment_paths = []):
    # create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = FROM
    message["To"] = to
    message["Subject"] = subject

    # add body
    message.attach(MIMEText(body, "plain"))

    # add attachments
    for ap in attachment_paths:
        attach_file_to_email(message, ap)
    
    text = message.as_string()

    return text

def attach_file_to_email(message, filepath):
    with open(filepath, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    filename = os.path.basename(filepath)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")

    message.attach(part)

if __name__ == "__main__":
    main()

