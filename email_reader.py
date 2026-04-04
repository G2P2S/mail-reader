import imaplib
import requests
from email import message_from_bytes
from email.header import decode_header
from bs4 import BeautifulSoup

def get_user_input():
    email = input("Email: ")
    client_id = input("Client id: ")
    refresh_token = input("Refresh token: ")
    return {
        "email": email,
        "client_id": client_id,
        "refresh_token": refresh_token
    }

class MailClient:
    def __init__(self, email, client_id, refresh_token):
        self.email = email
        self.client_id = client_id
        self.refresh_token = refresh_token
        self.access_token = None

    def get_access_token(self):
        data = {
            'client_id': self.client_id,
            'grant_type': 'refresh_token',
            'refresh_token': self.refresh_token
        }
        response = requests.post('https://login.live.com/oauth20_token.srf', data=data)
        response.raise_for_status()
        self.access_token = response.json()['access_token']

    def generate_auth_string(self, email, access_token):
        return f"user={email}\1auth=Bearer {access_token}\1\1"

    def decode_mime_words(self, s):
        decoded_fragments = []
        for word, enc in decode_header(s):
            if isinstance(word, bytes):
                decoded_fragments.append(word.decode(enc or 'utf-8', errors='replace'))
            else:
                decoded_fragments.append(word)
        return ''.join(decoded_fragments)

    def get_body_from_msg(self, msg):
        """Рекурсивно получаем все текстовые части письма"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                dispo = str(part.get("Content-Disposition") or "")
                if ctype.startswith("text/") and "attachment" not in dispo:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or 'utf-8'
                        try:
                            text = payload.decode(charset, errors='replace')
                        except Exception:
                            text = payload.decode('utf-8', errors='replace')
                        if ctype == "text/html":
                            text = BeautifulSoup(text, "html.parser").get_text()
                        body += text + "\n"
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or 'utf-8'
                try:
                    body = payload.decode(charset, errors='replace')
                except Exception:
                    body = payload.decode('utf-8', errors='replace')
                if msg.get_content_type() == "text/html":
                    body = BeautifulSoup(body, "html.parser").get_text()
        return body.strip()

    def connect_imap(self):
        mail = imaplib.IMAP4_SSL('outlook.office365.com')
        mail.authenticate('XOAUTH2', lambda x: self.generate_auth_string(self.email, self.access_token))
        mail.select("INBOX")

        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        print(f"Found {len(email_ids)} emails.\n")

        for i, num in enumerate(email_ids, start=1):
            status, data = mail.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            msg = message_from_bytes(raw_email)

            subject = self.decode_mime_words(msg.get("Subject", ""))
            from_ = self.decode_mime_words(msg.get("From", ""))

            body = self.get_body_from_msg(msg)

            print(f"--- Email {i} ---")
            print(f"From: {from_}")
            print(f"Subject: {subject}")
            print("Body:")
            print(body)
            print("="*60)

        mail.logout()

def run():
    user_data = get_user_input()
    client = MailClient(
        email=user_data['email'],
        client_id=user_data['client_id'],
        refresh_token=user_data['refresh_token']
    )
    client.get_access_token()
    client.connect_imap()

if __name__ == "__main__":
    run()