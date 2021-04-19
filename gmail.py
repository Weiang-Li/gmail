import imaplib
import email
import os
from dotenv import load_dotenv
from pathlib import Path

env_path = Path('.') / 'credentials.env'
load_dotenv(dotenv_path=env_path)

email_user = os.environ.get('my_email')
email_pwd = os.environ.get('my_email_password')


class Email:
    def __init__(self, email_username, email_password):
        self.email_username = email_username
        self.email_password = email_password
        self.con = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        self.con.login(self.email_username, self.email_password)
        self.con.select('INBOX')

    def get_email_date(self, email_subject, n_th_email):
        try:
            result, data = self.con.search(None, 'SUBJECT', email_subject)
            emails = data[0].split()[n_th_email]
            outcome, info = self.con.fetch(emails, '(RFC822)')
            raw_email_string = info[0][1].decode('utf-8')
            full_date_time = email.message_from_string(raw_email_string)['Date']
            email_date_string = str(full_date_time).split()
            date = '.'.join((email_date_string[2], email_date_string[1], email_date_string[3]))
            return date
        except Exception as e:
            print(e)

    def get_email_body(self, email_subject, n_th_email):
        lst = list(reversed(range(n_th_email, 0)))
        try:
            for number in lst:
                result, data = self.con.search(None, 'SUBJECT', email_subject)
                emails = data[0].split()[number]
                print(emails)
                outcome, info = self.con.fetch(emails, '(RFC822)')
                raw_email_string = info[0][1].decode('utf-8')
                return raw_email_string
        except Exception as e:
            print(e)

    def get_email_location(self, email_subject, n_th_email):
        lst = list(reversed(range(n_th_email, 0)))
        try:
            for number in lst:
                result, data = self.con.search(None, 'SUBJECT', email_subject)
                emails = data[0].split()[number]
                outcome, info = self.con.fetch(emails, '(RFC822)')
                raw_email_string = info[0][1].decode('utf-8')
                location = str(email.message_from_string(raw_email_string)['SUBJECT']).split()[1]
                return location
        except Exception as e:
            print(e)

    def get_emailattachment(self, email_subject, how_many_emails, output_folder):
        lst = list(reversed(range(how_many_emails, 0)))
        try:
            for number in lst:
                result, data = self.con.search(None, 'SUBJECT', email_subject)
                emails = data[0].split()[number]
                outcome, info = self.con.fetch(emails, '(RFC822)')
                downloadfile = email.message_from_bytes(info[0][1])
                print(email_subject, self.get_email_date(email_subject, number))
                for part in downloadfile.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    fileName = str(part.get_filename()).replace('\r\n', '')
                    if bool(fileName):
                        filePath = os.path.join(output_folder, fileName)
                        with open(filePath, 'wb') as f:
                            f.write(part.get_payload(decode=True))
        except Exception as e:
            print(e)


if __name__ == '__main__':
    # <-example->
    acc = Email('your email', 'your password')

    # <-get email attachment->
    # acc.get_emailattachment('TX-SA',how_many_emails=-1,output_folder=r'C:\Users\B\Desktop\Download From Email Automation')
    # <-get email attachment->

    # <-get email body->
    # print(acc.get_email_body('TX-SA',n_th_email=-5))    #n_th_email means: counting from the most recent one, how far back you want to go: in this case means the 5th most recent one
    # <-get email body->

    # <-get email location: first 5 characters from the subject->
    # print(acc.get_email_location('TX-SA',n_th_email=-5))
    # <-get email location: first 5 characters from the subject->

    # <-get email date->
    # print(acc.get_email_date('TX-SA', n_th_email=-1))
    # <-get email date->
