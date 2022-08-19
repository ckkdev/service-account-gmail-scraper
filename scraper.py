# importing os & os path & pathLib module
import os
import os.path
from pathlib import Path
from tkinter import E
import dateutil.parser as parser
from typing import List

# importing email module
from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from attachment import Attachment

# Parent Directory path
parent_dir = str(Path(os.getcwd())) + '\EMAILS'

#import BS & B64
import base64
from bs4 import BeautifulSoup


#importing gApi
from googleapiclient import discovery
from google.oauth2 import service_account

def AccountServiceScraper(email_users, account_scope, service_account_json):

    #loop through email users
    for email_user in email_users:
        credentials = service_account.Credentials.from_service_account_file(service_account_json, scopes=account_scope)
        delegated_credentials = credentials.with_subject(email_user)

        service = discovery.build('gmail', 'v1', credentials=delegated_credentials)
        user_path = os.path.join(parent_dir, email_user)
        attachments = 'reference'
        def evaluate_message_payload(
            payload: dict,
            user_id: str,
            msg_id: str
        ) ->List[dict]:
            """
            Recursively evaluates a message payload.
            Args:
                payload: The message payload object (response from Gmail API).
                user_id: The current account address (default 'me').
                msg_id: The id of the message.
                attachments: Accepted values are 'ignore' which completely ignores
                    all attachments, 'reference' which includes attachment
                    information but does not download the data, and 'download' which
                    downloads the attachment data to store locally. Default
                    'reference'.
            Returns:
                A list of message parts.
            Raises:
                googleapiclient.errors.HttpError: There was an error executing the
                    HTTP request.
            """

            if 'attachmentId' in payload['body']:  # if it's an attachment
                if attachments == 'ignore':
                    return []

                att_id = payload['body']['attachmentId']
                filename = payload['filename']
                if not filename:
                    filename = 'unknown'

                obj = {
                    'part_type': 'attachment',
                    'filetype': payload['mimeType'],
                    'filename': filename,
                    'attachment_id': att_id,
                    'data': None
                }

                if attachments == 'reference':
                    return [obj]

                else:  # attachments == 'download'
                    if 'data' in payload['body']:
                        data = payload['body']['data']
                    else:
                        res = service.users().messages().attachments().get(
                            userId=user_id, messageId=msg_id, id=att_id
                        ).execute()
                        data = res['data']

                    file_data = base64.urlsafe_b64decode(data)
                    obj['data'] = file_data
                    return [obj]

            elif payload['mimeType'] == 'text/html':
                data = payload['body']['data']
                data = base64.urlsafe_b64decode(data)
                body = BeautifulSoup(data, 'lxml', from_encoding='utf-8').body
                return [{ 'part_type': 'html', 'body': str(body) }]

            elif payload['mimeType'] == 'text/plain':
                data = payload['body']['data']
                data = base64.urlsafe_b64decode(data)
                body = data.decode('UTF-8')
                return [{ 'part_type': 'plain', 'body': body }]

            elif payload['mimeType'].startswith('multipart'):
                ret = []
                if 'parts' in payload:
                    for part in payload['parts']:
                        ret.extend(evaluate_message_payload(part, user_id, msg_id))
                return ret

            return []

        #create parent folder for user@email.com
        try: 
            os.makedirs(user_path) 
            print("Directory '% s' created" % user_path)
        except OSError as error: 
            print(error)

        res = service.users().labels().list(userId=email_user).execute()
        labels = res['labels']
        for label in labels:

            path = user_path + '/' + label['name']
            directory_path = os.path.join(parent_dir, path)
            path = os.path.join(parent_dir, directory_path)

            #create folders for each label in the users mail
            try: 
                os.makedirs(path) 
                print("Directory '% s' created" % path)
            except OSError as error: 
                print(error)

            #get all messages for each folder
            response = service.users().messages().list(
                userId=email_user,
                labelIds=[label['id']],
                includeSpamTrash=True
            ).execute()

            message_refs = []
            if 'messages' in response:  # ensure request was successful
                message_refs.extend(response['messages'])

            while 'nextPageToken' in response:
                
                page_token = response['nextPageToken']
                response = service.users().messages().list(
                userId=email_user,
                labelIds=[label['id']],
                includeSpamTrash=True,
                pageToken=page_token
                ).execute()
                message_refs.extend(response['messages'])
                
            #loop through messages
            for ref in message_refs:
                
                message = service.users().messages().get(
                    userId=email_user, id=ref['id']
                ).execute()

                payload = message['payload']
                headers = payload['headers']

                msg_hdrs = {}
                for hdr in headers:
                    if hdr['name'].lower() == 'date':
                        try:
                            date = str(parser.parse(hdr['value']).astimezone())
                        except Exception:
                            date = hdr['value']
                    elif hdr['name'].lower() == 'from':
                        sender = hdr['value']
                    elif hdr['name'].lower() == 'to':
                        recipient = hdr['value']
                    elif hdr['name'].lower() == 'subject':
                        subject = hdr['value']

                    msg_hdrs[hdr['name']] = hdr['value']

                plain_msg = None
                html_msg = None
                
                parts = evaluate_message_payload(
                    payload, email_user, ref['id']
                )

                plain_msg = None
                html_msg = None
                attms = []
                for part in parts:
                    if part['part_type'] == 'plain':
                        if plain_msg is None:
                            plain_msg = part['body']
                        else:
                            plain_msg += '\n' + part['body']
                    elif part['part_type'] == 'html':
                        if html_msg is None:
                            html_msg = part['body']
                        else:
                            html_msg += '<br/>' + part['body']
                    elif part['part_type'] == 'attachment':
                        attm = Attachment(service, email_user, ref['id'],
                                        part['attachment_id'], part['filename'],
                                        part['filetype'], part['data'])
                        attms.append(attm)
                #setting up email & duplicating headers + payload
                msg = MIMEMultipart('alternative')
                msg.set_charset("utf-8")
                index = 1
                headers = msg_hdrs
                for header in headers:
                    msg[header] = headers[header]
                    index += 1

                if html_msg is not None:
                    body_content = MIMEText(html_msg, 'html')
                else:
                    body_content = MIMEText(plain_msg, 'plain')

                msg.attach(body_content)
                #creating .eml with email ID as the name
                outfile_name = os.path.join(parent_dir, directory_path, ref['id']+".eml")
                print("Duplicated email " + ref['id'] + " into folder " + label['name'])
                with open(outfile_name, 'w') as outfile:
                    gen = generator.Generator(outfile)
                    gen.flatten(msg)


