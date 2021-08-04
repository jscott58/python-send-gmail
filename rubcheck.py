import pysftp
import os
import sys
import time
import base64
import httplib2
from io import StringIO
import oauth2client
from oauth2client import client, tools, file
from email.mime.text import MIMEText
from apiclient import errors, discovery
import smtplib

astline = "************************************************************************"
sender = "scottcalhoun65@gmail.com"
to = "scott.calhoun@hbkeso.com"
subject = "Utilization Run Check"
message_texts = [""]
userid = "scottcalhoun65@gmail.com"

SCOPES = "https://www.googleapis.com/auth/gmail.send"
CLIENT_SECRET_FILE = ".\client_secret_554972950047-tgika10cfs7klt21kb2ototpfa87cqpf.apps.googleusercontent.com.json"
APPLICATION_NAME = "GmailRunCHeck"


def resource_path(relative_path):
    # Get absolute path to resource, works for dev and for PyInstaller
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def get_credentials():
    home_dir = os.path.expanduser("~")
    credential_dir = os.path.join(home_dir, ".credentials")
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, "gmail-python-email-send.json")
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print("Storing credentials to " + credential_path)
    return credentials


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
    message = MIMEText(message_text)
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
    except errors.HttpError as exception:
        print("An error occurred: %s" % message.HttpError)


credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
service = discovery.build("gmail", "v1", http=http)

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
# ASpecify host jey for sftp server
cnopts = pysftp.CnOpts(knownhosts=sftphostkey_path)
# And authenticate with a private key
with pysftp.Connection(
    host="sftp.hbkeso.com",
    username="HBKUtil",
    password="^8y5Sc1F3%e89xZ2#blkW6pn0(Ym)R@s1!G",
    cnopts=cnopts,
) as sftp:
    try:
        utilFiles = sftp.listdir("./")  # upload file to public/ on remote
        print(astline)
        print("")
        if len(utilFiles) == 0:
            print("No files found.")
        else:
            print("sftp files:")
            # printing the list using loop
            for x in range(len(utilFiles)):
                print("      " + utilFiles[x])
                fileStatsObj = os.stat(utilFiles[x])
                modificationTime = time.ctime(fileStatsObj.st_mtime)
                print("      Last Modified: ", modificationTime)
                message_texts[x] = utilFiles[x] + " Last Modified: " + modificationTime
                emailObject = create_message(sender, to, subject, message_texts[x])
                returnMessage = send_message(service, userid, emailObject)
        print("")
        print(astline)
    except ValueError:
        print("SFTP listdir ERROR - EXITING")
        print(astline)
