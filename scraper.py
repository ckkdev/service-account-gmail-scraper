# importing os & os path & pathLib module
import os
import os.path
from pathlib import Path
import dateutil.parser as parser

# importing email module
from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Parent Directory path
parent_dir = str(Path(os.getcwd())) + '\EMAILS'

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
                for part in payload['body']:
                    if part[1] == 'plain':
                        if plain_msg is None:
                            plain_msg = part['body']
                        else:
                            plain_msg += '\n' + part['body']
                    elif part[0] == 'html':
                        if html_msg is None:
                            html_msg = part['body']
                        else:
                            html_msg += '<br/>' + part['body']

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

                elif plain_msg is not None:
                    body_content = MIMEText(plain_msg, 'plain')

                else:
                    body_content = MIMEText('n/a', 'plain')

                msg.attach(body_content)

                #creating .eml with email ID as the name
                outfile_name = os.path.join(parent_dir, directory_path, ref['id']+".eml")
                print("Duplicated email " + ref['id'] + " into folder " + label['name'])
                with open(outfile_name, 'w') as outfile:
                    gen = generator.Generator(outfile)
                    gen.flatten(msg)