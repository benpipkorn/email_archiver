import os
import pickle

# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode

# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = 'benpipkorn@gmail.com'
cur_credentials = "benpipkorn_credentials.json"

def gmail_authenticate():
    creds = None
    # load a set of access and refresh tokens already in use
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no credentials or they are invalid, allow login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cur_credentials, SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for next run, so as not to make user login every time
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# preparing and adding an attachement to a message
def add_attachement(message, file_name):
    content_type, encoding = guess_mime_type(file_name)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(file_name, 'rb')
        msg = MIMEText(fp.read().decode(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(file_name, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(file_name, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else: 
        fp = open(file_name, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    file_name = os.path.basename(file_name)
    msg.add_header('Content-Disposition', 'attachement', filename=file_name)
    message.attach(msg)

# building a message to send
def build_message(destination, obj, body, attachments=[]):
    if not attachments:
        message = MIMEText(body)
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
    else:
        message = MIMEMultipart()
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
        message.attach(MIMEText(body))
        for filename in attachments:
            add_attachement(message, filename)
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

# compile and send a message
def send_message(service, destination, obj, body, attachments=[]):
    return service.users().messages().send(
        userId = 'me',
        body = build_message(destination, obj, body, attachments)
    ).execute()

# search for emails using a specific query
def search(service, query):
    result = service.users().messages().list(userId = 'me', q = query).execute()
    messages = []
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId = 'me', q = query, pageToken = page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
        if len(messages) >= 1000:
            break
    return messages

# a function specifically to find unread emails
def get_unread(service):
    result = service.users().messages().list(userId = 'me', labelIds = ['INBOX','UNREAD']).execute()
    messages = []
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId = 'me', labelIds = ['UNREAD'], pageToken = page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
        if len(messages) >= 1000:
            break
    return messages


# delete messages with a given query
def delete_messages(service, messages):
    if len(messages) < 1:
        return
    return service.users().messages().batchModify(
        userId = 'me',
        body = {
            'ids': [msg['id'] for msg in messages],
            'addLabelIds': ['TRASH'],
            'removeLabelIds': ['UNREAD']
        }
    ).execute()

def form_summary(messages, file_name):
    with open(file_name, 'w') as summary:
        summary.write('Total emails deleted: ' + str(len(messages)) + '\n')
        counter = 1
        for message in messages:
            sender = ''
            subject = ''
            msg = service.users().messages().get(userId = 'me', id = message['id'], format = 'full').execute()
            payload = msg['payload']
            headers = payload.get('headers')
            if headers:
                for header in headers:
                    name = header.get('name').lower()
                    value = header.get('value')
                    if name == 'from':
                        sender = ascii(value)
                    if name == 'subject':
                        subject = ascii(value)
            summary.write("\n---\n")
            summary.write(str(counter) + " : " + sender + " : " + subject)
            counter += 1
    return
    


if __name__ == "__main__":
    # get Gmail API service
    service = gmail_authenticate()

    # searching for unread emails
    unread_messages = get_unread(service)

    # creating and sending a summary of the emails moved to trash
    form_summary(unread_messages, 'summary.txt')
    with open('summary.txt', 'r') as summary:
        send_message(service, our_email, 'Summary of Deleted Emails', summary.read())

    # deleting all unread emails
    delete_messages(service, unread_messages)

    # test email
    # send_message(service, our_email, 'Test Email', 'Test Body')

