from __future__ import print_function
import pysftp
import os
import sys
import time
import datetime
import base64
import glob
from email.mime.text import MIMEText
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 
          'https://www.googleapis.com/auth/gmail.send']

astline = "***********************************************************************"
gmail_sender = "scottcalhoun65@gmail.com"
gmail_to = "scott.calhoun@hbkeso.com"
gmail_subject = "Utilization Run Check"
gmail_userid = "scottcalhoun65@gmail.com"
sftp_host = "sftp.hbkeso.com"
sftp_username = "HBKUtil"
sftp_password = "^8y5Sc1F3%e89xZ2#blkW6pn0(Ym)R@s1!G"


def resource_path(relative_path):
    # Get absolute path to resource, works for dev and for PyInstaller
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def create_message(sender, to, subject, message_text):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      message_text: The text of the email message.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEText(message_text, _charset='UTF-8')
    message["to"] = to
    message["from"] = sender
    message["subject"] = subject
    # bytes-like object is required, not 'str' Error on origianl code
    # return {"raw": base64.urlsafe_b64encode(message.as_string())}
    return {"raw": base64.urlsafe_b64encode(message.as_string().encode()).decode()}


def send_message(service, user_id, message):
    """Send an email message.

    Args:
      service: Authorized Gmail API service instance.
      user_id: User's email address. The special value "me"
      can be used to indicate the authenticated user.
      message: Message to be sent.

    Returns:
      Sent Message.
    """
    try:
        message = (
            service.users().messages().send(userId=user_id, body=message).execute()
        )
        print("Message Id: %s" % message["id"])
        return message
    except service.HttpError as exception:
        print("An error occurred: %s" % message.HttpError)
        

def empty_accountroot(folder_to_empty):
    for CleanUp in glob.glob(folder_to_empty + '/*.*'):
        if not CleanUp.endswith('token.json'):    
            os.remove(CleanUp)


def main():
    gmail_text = ""
    local_dir = ".\\accountroot"
    # Setup local accountroot folder if not eexisit
    if os.path.isdir(local_dir):
        empty_accountroot(local_dir)
    else:
        os.mkdir(local_dir)
    print(astline)
    print("HBK Solutions LLC Unanet Utilization Report - Run Check")
    print(astline)
    print("")

    print(astline)
    print("Reading HBKUtil utilization folder on sftp.hbkeso.com")
    print(astline)
    print("")
    # Define hostkey path so that pyinstaller will correctly include when using --onrfile
    sftphostkey_path = resource_path("hbkeso.pub")
    # A Specify host jey for sftp server
    cnopts = pysftp.CnOpts(knownhosts=sftphostkey_path)
    # And authenticate with a private key
    with pysftp.Connection(
        host=sftp_host,
        username=sftp_username,
        password=sftp_password,
        cnopts=cnopts,
    ) as sftp:
        try:
            utilFiles = sftp.listdir("./")  # upload file to public/ on remote
            print(astline)
            if len(utilFiles) == 0:
                print("No files found.")
            else:
                print("sftp files:")
                sftp.get_d('.', local_dir, preserve_mtime=True)
                # printing the list using loop
                for f in os.listdir(local_dir):
                    print(f)
                    utilFileStatsObj = os.stat(os.path.join(local_dir, f))
                    modificationTime = time.ctime(utilFileStatsObj.st_mtime)
                    print("    Last Modified: ", modificationTime)
                    gmail_text += f + "\n" + "      Last Modified: " + modificationTime + "\n\r" 
            print(astline)
            print("")
            try:
                creds = None
                # The file token.json stores the user's access and refresh tokens, and is
                # created automatically when the authorization flow completes for the first
                # time.
                if os.path.exists(os.path.join(local_dir,'token.json')):
                    creds = Credentials.from_authorized_user_file(os.path.join(local_dir,'token.json'), SCOPES)
                # If there are no (valid) credentials available, let the user log in.
                if not creds or not creds.valid:
                    if creds and creds.expired and creds.refresh_token:
                        creds.refresh(Request())
                    else:
                        flow = InstalledAppFlow.from_client_secrets_file(
                            resource_path('credentials.json'), SCOPES)
                        creds = flow.run_local_server(port=0)
                        # creds = ServiceAccountCredentials.from_json_keyfile_name(resource_path('credentials.json'), SCOPES)
                    # Save the credentials for the next run
                    with open(os.path.join(local_dir,'token.json'), 'w') as token:
                        token.write(creds.to_json())

                service = build('gmail', 'v1', credentials=creds)

                # Call the Gmail API
                results = service.users().labels().list(userId='me').execute()
                labels = results.get('labels', [])
                print(astline)
                if not labels:
                    print('No labels found.')
                else:
                    print('Gmail Account Found')
                    # enable label print to debug access to gmail account
                    # print('Labels:')
                    # for label in labels:
                    #     print(label['name'])
                        
                    gmail_message = create_message(gmail_sender, gmail_to, gmail_subject, gmail_text)
                    send_message(service, gmail_userid, gmail_message)
                    # print(gmail_text)
                print(astline)
            except ValueError:
                    print("Google API ERROR: " + ValueError)
        except ValueError:
            print("SFTP ERROR - EXITING: " + ValueError)
            print(astline)


if __name__ == '__main__':
    main()