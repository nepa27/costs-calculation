from email import message_from_string
from imaplib import IMAP4_SSL


def get_messages_from_email():
    """ Get messages from email. """
    messages = []
    mail = IMAP4_SSL('imap.yandex.ru')  # write your email
    with open('pass.txt', 'r') as line:
        _ = line.readlines()
        login = _[0].split('login:')
        passw = _[1].split('pass:')
    mail.login(f'{login[1][:-1]}', f'{passw[1]}')

    mail.list()
    mail.select("inbox")

    result, data = mail.search(None, "ALL")

    ids = data[0]
    ids = ids.split()
    ids.reverse()
    latest_email = ids[:10]
    for latest_email_id in latest_email:
        result, data = mail.fetch(latest_email_id, "(RFC822)")
        raw_email = data[0][1]
        raw_email_string = raw_email.decode('utf-8')

        email_message = message_from_string(raw_email_string)
        try:
            if email_message.is_multipart():
                for payload in email_message.get_payload():
                    body = payload.get_payload(decode=True).decode('utf-8')
                    messages.append(body)
            body = email_message.get_payload(decode=True).decode('utf-8')
            messages.append(body)
        except AttributeError:
            with open('log.txt', 'a') as error:
                error.write('Didn`t decode message\n')

    return messages
