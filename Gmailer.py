import httplib2
import os
import oauth2client
from oauth2client import client, tools
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apiclient import errors, discovery
from datetime import datetime

# ToDo to config file
SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'C:\ProgramData\MathnasiumScheduler\Working Directory\client_secret.1.json'
APPLICATION_NAME = 'Gmail API Python Send Email'


class Gmailer:

    @staticmethod
    def get_credentials():
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir, 'gmail-python-email-send.json')
        store = oauth2client.file.Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            credentials = tools.run_flow(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    @classmethod
    def send_message(cls, sender, to, subject, msgHtml, msgPlain):
        credentials = cls.get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('gmail', 'v1', http=http)
        message = cls.create_message(sender, to, subject, msgHtml, msgPlain)
        cls.send_message_internal(service, "me", message)

    @staticmethod
    def send_message_internal(service, user_id, message):
        try:
            message = (service.users().messages().send(userId=user_id, body=message).execute())
            print('Message Id: %s' % message['id'])
            return message
        except errors.HttpError as error:
            print('An error occurred: %s' % error)

    @staticmethod
    def create_message(sender, to, subject, msgHtml, msgPlain):
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = to
        msg.attach(MIMEText(msgPlain, 'plain'))
        msg.attach(MIMEText(msgHtml, 'html'))
        raw = base64.urlsafe_b64encode(msg.as_bytes())
        raw = raw.decode()
        body = {'raw': raw}
        return body

    @classmethod
    def send_instructor_schedule(cls, instructor):
        test_instructor_address = 'stafford@mathnasium.com' # ToDo for test purposes only, comment out when testing ends
        center_email_address = 'stafford@mathnasium.com'  # or 'chancellor@email.com'
        subject = 'Mathnasium Work Schedule effective ' + str(datetime.now().today().date())
        # cls.send_message(center_email_address, instructor.email_address, instructor.schedule_as_html_msg(),
        #              instructor.schedule_as_plain_msg())
        cls.send_message(center_email_address, test_instructor_address, subject, instructor.schedule_as_html_msg(),
                         instructor.schedule_as_plain_msg())
    @classmethod
    def send_instructor_schedules(cls, instructors):
        for each_instructor in instructors: cls.send_instructor_schedule(each_instructor)

#  def main():
#     to = "gerald.depasquale@gmail.com" # ""to@address.com"
#     sender = "gerald.depasquale@gmail.com" # "from@address.com"
#     subject = "Email from Scheduler App" # "subject"
#     msgHtml = "Dear Gerald,<br/>Awesome! Gmail sent successfully!<br/>Incredible,<br/>Jerry "# "Hi<br/>Html Email"
#     msgPlain = "\n" #"Hi\nPlain Email"
#     send_message(sender, to, subject, msgHtml, msgPlain)
#
# if __name__ == '__main__':
#     main()