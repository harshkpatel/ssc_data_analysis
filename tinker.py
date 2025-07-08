import imaplib
import email
from pyexpat.errors import messages


def main():
   username = input('Enter your username: ')
   password = input('Enter your password: ')
   imap_server = imaplib.IMAP4_SSL('outlook.office.com')
   imap_server.login(username, password)
   messages = imap_server.select('INBOX')

if __name__ == '__main__':
    main()