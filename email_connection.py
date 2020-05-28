#!/usr/bin/python3
import typing
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl
import os


class EmailSender:
    def __init__(self, smtp_server, smtp_port=25, username=None, password=None):
        self.smtpAddress = smtp_server
        self.smtpPort = smtp_port
        self.username = username
        self.password = password

    def send_email(self, from_address: str, to_addresses: typing.List[str], subject: str = '', body: str = '',
                   attachments_paths: typing.List[str] = None):
        try:
            message = MIMEMultipart()
            message['From'] = from_address
            message['To'] = ', '.join(to_addresses)
            message['Subject'] = subject
            message.attach(MIMEText(body, "plain"))
            if attachments_paths:
                for attachment_path in attachments_paths:
                    with open(attachment_path, 'rb') as atch:
                        # Add file as application/octet-stream
                        # Email client can usually download this automatically as attachment
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(atch.read())
                    # Encode file in ASCII characters to send by email
                    encoders.encode_base64(part)
                    # Add header as key/value pair to attachment part
                    filename = os.path.basename(attachment_path)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename= {filename}",
                    )
                    # Add attachment to message and convert message to string
                    message.attach(part)
            text = message.as_string()
            if self.username and self.password:
                # Log in to server using secure context and send email
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(host=self.smtpAddress, port=self.smtpPort, context=context) as server:
                    server.login(self.username, self.password)
                    server.sendmail(from_address, to_addresses, text)
            else:
                with smtplib.SMTP(host=self.smtpAddress, port=self.smtpPort) as server:
                    server.sendmail(from_address, to_addresses, text)
            return True
        except Exception as e:
            print(e)
            return False