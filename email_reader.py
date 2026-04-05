import base64
import imaplib
import requests
import email
from email.header import decode_header
import webbrowser
import os

imap_server = 'outlook.office365.com'


def get_user_input():
    email_addr = input("Email: ")
    client_id = input("Client id: ")
    refresh_token = input("Refresh token: ")
    return {
        "email": email_addr,
        "client_id": client_id,
        "refresh_token": refresh_token
    }


class MailClient:

    def __init__(self, email_addr, client_id, refresh_token):
        self.email = email_addr
        self.client_id = client_id
        self.refresh_token = refresh_token
        self.access_token = None

    def get_access_token(self):
        data = {
            'client_id': self.client_id,
            'grant_type': 'refresh_token',
            'refresh_token': self.refresh_token
        }

        response = requests.post(
            'https://login.live.com/oauth20_token.srf',
            data=data
        )

        if response.status_code != 200:
            raise Exception(f"Ошибка токена: {response.text}")

        self.access_token = response.json()['access_token']

    def generate_auth_string(self):
        return f"user={self.email}\1auth=Bearer {self.access_token}\1\1"

    def decode_mime_words(self, s):
        decoded = decode_header(s)
        result = ""
        for part, encoding in decoded:
            if isinstance(part, bytes):
                result += part.decode(encoding or "utf-8", errors="ignore")
            else:
                result += part
        return result

    def extract_body(self, msg):
        html = None
        text = None

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" in content_disposition:
                    continue

                payload = part.get_payload(decode=True)
                if not payload:
                    continue

                charset = part.get_content_charset() or "utf-8"

                try:
                    content = payload.decode(charset, errors="ignore")
                except:
                    content = payload.decode("utf-8", errors="ignore")

                if content_type == "text/html":
                    html = content
                elif content_type == "text/plain":
                    text = content
        else:
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or "utf-8"
            html = payload.decode(charset, errors="ignore")

        return html if html else f"<pre>{text}</pre>"

    def connect_imap(self):
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.authenticate('XOAUTH2', lambda x: self.generate_auth_string())
        mail.select("INBOX")

        status, messages = mail.search(None, 'ALL')
        mail_ids = messages[0].split()

        print(f"Found {len(mail_ids)} emails")

        with open("emails.html", "w", encoding="utf-8") as f:
            f.write("""
            <html>
            <head>
                <meta charset="UTF-8">
                <title>Emails</title>
            </head>
            <body>
            <h1>Мои письма</h1>
            <hr>
            """)

            for i, mail_id in enumerate(mail_ids, start=1):
                status, msg_data = mail.fetch(mail_id, "(RFC822)")
                raw_email = msg_data[0][1]

                msg = email.message_from_bytes(raw_email)

                subject = self.decode_mime_words(msg.get("Subject", "No subject"))
                from_ = self.decode_mime_words(msg.get("From", "Unknown"))

                body = self.extract_body(msg)

                f.write(f"<h2>{i}. {subject}</h2>")
                f.write(f"<p><b>From:</b> {from_}</p>")
                f.write("<hr>")
                f.write(body)
                f.write("<br><br><hr><hr>")

                print(f"Processed email {i}")

            f.write("""
            </body>
            </html>
            """)

        print("All emails saved to emails.html")

        file_path = os.path.abspath("emails.html")
        webbrowser.open(f"file://{file_path}")

        mail.logout()


def run():
    user_data = get_user_input()

    client = MailClient(
        email_addr=user_data['email'],
        client_id=user_data['client_id'],
        refresh_token=user_data['refresh_token']
    )

    client.get_access_token()
    client.connect_imap()


if __name__ == "__main__":
    run()