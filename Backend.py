# Importing required packages 
import httplib2, os, oauth2client, base64, mimetypes, json, pytz, openai
from datetime import timedelta
from datetime import datetime as dt
import requests, time, logging
from oauth2client import client, tools, file
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apiclient import errors, discovery
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import dotenv
from dotenv import load_dotenv

# Load environment variables -- the .env file contains important settings for the
# application, such as the user's email, timezone, and OpenAI API key
dotenv_path = os.path.abspath(".env")
load_dotenv(dotenv_path)

# High-level parameters
user_email = os.getenv("USER_EMAIL")
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/calendar",
]
GOOGLE_CLIENT_SCRET_FILE = "google_client_secret.json"
GMAIL_APPLICATION_NAME = "Email Automation"

# Setting the API key for OpenAI
openai.api_key = os.getenv("OPENAI_API_KEY")

# User's Gmail credentials
home_dir = os.path.expanduser("~")
credential_dir = os.path.join(home_dir, ".credentials")

# Credential directory
if not os.path.exists(credential_dir):
    os.makedirs(credential_dir)

# Path to user's credentials
split_email = user_email.split("@")
split_email[1] = split_email[1].replace(".", "-")
short_path = split_email[0] + "-" + split_email[1] + "-"
credential_path = os.path.join(credential_dir, short_path + "google-python-email.json")


# Function to get user's gmail credentials
def get_google_credentials():
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(GOOGLE_CLIENT_SCRET_FILE, GOOGLE_SCOPES)
        flow.user_agent = GMAIL_APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
    return credentials


# Function to send gmail message
def gmail_send(sender, to, subject, msgHtml, attachmentFile=None):
    credentials = get_google_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build("gmail", "v1", http=http)
    if attachmentFile:
        message1 = create_gmail_with_attachment(
            sender, to, subject, msgHtml, attachmentFile
        )
    else:
        message1 = create_gmail_html(sender, to, subject, msgHtml)
    result = send_gmail_internal(service, "me", message1)
    return result


# Helper function to send gmail message
def send_gmail_internal(service, user_id, message):
    try:
        message = (
            service.users().messages().send(userId=user_id, body=message).execute()
        )
        return message
    except errors.HttpError as error:
        print("An error occurred: %s" % error)
        return "Error"


# Function to create gmail message HTML
def create_gmail_html(sender, to, subject, msgHtml):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = to
    msg.attach(MIMEText(msgHtml, "html"))
    return {"raw": base64.urlsafe_b64encode(msg.as_string().encode()).decode()}


# Function to create gmail message with the possbility of including an attachment
def create_gmail_with_attachment(sender, to, subject, msgHtml, attachmentFile):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      msgHtml: Html message to be sent
      attachmentFile: The path to the file to be attached.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEMultipart("mixed")
    message["to"] = to
    message["from"] = sender
    message["subject"] = subject

    messageA = MIMEMultipart("alternative")
    messageR = MIMEMultipart("related")

    messageR.attach(MIMEText(msgHtml, "html"))
    messageA.attach(messageR)

    message.attach(messageA)

    print("create_message_with_attachment: file: %s" % attachmentFile)
    content_type, encoding = mimetypes.guess_type(attachmentFile)

    # Checking for the content type of the attachment
    if content_type is None or encoding is not None:
        content_type = "application/octet-stream"
    main_type, sub_type = content_type.split("/", 1)

    # Checking for the main type of the attachment
    if main_type == "text":
        fp = open(attachmentFile, "rb")
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == "image":
        fp = open(attachmentFile, "rb")
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == "audio":
        fp = open(attachmentFile, "rb")
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(attachmentFile, "rb")
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()

    # Encoding the attachment
    filename = os.path.basename(attachmentFile)
    msg.add_header("Content-Disposition", "attachment", filename=filename)
    message.attach(msg)

    return {"raw": base64.urlsafe_b64encode(message.as_string())}

# Function to read gmail messages
def gmail_read():
    # The file gmail_token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = get_google_credentials()
    output = []
    try:
        # Call the Gmail API
        service = build("gmail", "v1", credentials=creds)
        results = (
            service.users()
            .messages()
            .list(userId="me", labelIds=["INBOX"], q="is:unread")
            .execute()
        )
        messages = results.get("messages", [])

        # If there are no unread messages, return a notification message
        if not messages:
            return ["No new messages."]
        else:

            # If there are unread messages, iterate through them 
            # and return the sender and the message body
            for message in messages:
                msg = (
                    service.users()
                    .messages()
                    .get(userId="me", id=message["id"])
                    .execute()
                )
                email_data = msg["payload"]["headers"]

                # Iterate through the headers to get the sender's name and email
                for values in email_data:
                    name = values["name"]
                    if name == "From":
                        from_name = values["value"]
                        for part in msg["payload"]["parts"]:
                            try:
                                data = part["body"]["data"]
                                byte_code = base64.urlsafe_b64decode(data)
                                text = str(byte_code.decode("utf-8"))
                                output.append((from_name, text))

                                # mark the message as read 
                                msg = (
                                    service.users()
                                    .messages()
                                    .modify(
                                        userId="me",
                                        id=message["id"],
                                        body={"removeLabelIds": ["UNREAD"]},
                                    )
                                    .execute()
                                )

                                return output
                            except BaseException as error:
                                pass

    # If there's an error, print it
    except Exception as error:
        print(f"An error occurred: {error}")

# Function to call the freebusy() function of the Google Calendar API
def check_schedule(start_time, end_time):
    creds = get_google_credentials()
    try:
        # Call the Google Calendar API
        service = build("calendar", "v3", credentials=creds)

        # Find next calendar availability between start and end times
        events_result = (
            service.freebusy()
            .query(
                body={
                    "timeMin": start_time,
                    "timeMax": end_time,
                    "timeZone": "UTC",
                    "items": [{"id": "primary"}],
                }
            )
            .execute()
        )
        events = events_result.get("calendars", [])
        return (events, events_result)

    except Exception as error:
        raise error


# Function to read Google Calendar events
def read_gcal_events(start_time, end_time, meeting_duration):
    try:
        # Call the Google Calendar API
        now = dt.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time

        # Set end interval to four weeks from now
        end_interval = (dt.utcnow() + timedelta(days=30)).isoformat() + "Z"

        # Duration in datetime format
        event_duration = timedelta(minutes=meeting_duration)

        # Call the check_schedule() function
        (events, events_result) = check_schedule(now, end_interval)

        # Output array
        non_avai = []

        # No events found
        if not events:
            return non_avai

        # iterating through events
        for event in events_result["calendars"]["primary"]["busy"]:
            # Convert event start and end times to datetime format
            event_start = dt.strptime(event["start"][:-1], "%Y-%m-%dT%H:%M:%S")
            event_end = event_start + event_duration

            # Formatting
            start_formatted = event_start.isoformat() + "Z" 
            end_formatted = event_end.isoformat() + "Z"

            # Check if event fits within time constraints
            if (
                event_start.time() >= dt.strptime(start_time, "%H:%M").time()
                and event_end.time() <= dt.strptime(end_time, "%H:%M").time()
            ):
                # Check if event truly fits within time constraints
                if len(check_schedule(start_formatted, end_formatted)[1]) > 0:
                    # Add event to output array
                    non_avai.append((event_start, event_end))

        return non_avai

    except Exception as error:
        raise error


# Function to create Google Calendar events
def create_gcal_event(speaker_name, start_time, end_time):

    creds = get_google_credentials()

    # Call the Google Calendar API
    service = build("calendar", "v3", credentials=creds)

    # Define the event details
    event = {
        'summary': 'Meeting with ' + speaker_name,
        'location': 'Virtual',
        'description': 'Meeting scheduled via email with love by your friendly neighborhood AI assistant :).',
        'start': {
            'dateTime': start_time.isoformat(),
            'timeZone': os.getenv('TIMEZONE'),
        },
        'end': {
            'dateTime': end_time.isoformat(),
            'timeZone': os.getenv('TIMEZONE'),
        },
        'reminders': {
            'useDefault': True,
        },
    }

    # Call the API to insert the event
    try:
        event = service.events().insert(calendarId="primary", body=event).execute()
        return event.get('htmlLink')
    except Exception as error:
        print('An error occurred: %s' % error)

# Function to schedule a meeting
def schedule_meeting(start_time, end_time, meeting_duration, speaker):

    # Call function to read a Google Calendar event
    busy_times = read_gcal_events(start_time, end_time, meeting_duration)

    # Find non-busy time within the time constraints
    free_time = []

    # In the case that busy_times is empty or before the first slot, 
    # schedule the meeting on the first available time slot
    if len(busy_times) == 0 or busy_times[0][0][0] > start_time + timedelta(minutes=meeting_duration):
        free_time = [start_time, start_time + timedelta(minutes=meeting_duration)]

    # Else, iterate through busy_times to find free times
    else:
        for index, event in enumerate(busy_times):
            # Start condition checking
            start_condition = False
            suggested_start = event[0][1]
            if suggested_start >= start_time:
                if index < len(busy_times) - 1 and suggested_start < busy_times[index + 1][0][0]:
                    start_condition = True
                elif index == len(busy_times) - 1 and suggested_start < end_time:
                    start_condition = True

            # End condition checking
            suggested_end = suggested_start+ meeting_duration
            if start_condition and suggested_end <= end_time:
                if index < len(busy_times) - 1 and suggested_end < busy_times[index + 1][0][0]:
                    free_time.append((suggested_start, suggested_end))
                else:
                    free_time.append((suggested_start, suggested_end))

    # No free times found
    if len(free_time) == 0:
        print("No free time found.")
        return None

    else:
        # Create a Google Calendar event
        return create_gcal_event(speaker, free_time[0][0], free_time[0][1])

# Ask OpenAI to generate a response
def ask_openai(email_text):

    # Set up API request parameters
    model_engine = "gpt-4"
    prompt_text = f"""If the following email wants to schedule a meeting, 
    output 'SCHEDULE_MT' and nothing else. Otherwise, response to 
    the following email as if you were {os.getenv("USER_NAME")}. Email text:\n{email_text}"""
    conversation = [{"role": "system", "content": prompt_text}]

    # Generate response using GPT-3 API
    response = openai.ChatCompletion.create(
        model=model_engine, 
        messages=conversation,
        temperature=0.5
    )

    # Print generated response
    return response.choices[0].message.content

# Main function
def main():

    # Time bounds
    start_time = os.getenv("START_WINDOW")
    end_time = os.getenv("END_WINDOW")

    # Meeting duration parameter
    meeting_duration = os.getenv("MEETING_DURATION")

    # Read emails
    unread_messages = gmail_read()
    
    # Analyze emails
    for message in unread_messages:

        # Ask OpenAI to generate a response
        response = ask_openai(message[1])

        # Splitting sender identification 
        split_sender = message[0].split(" ")

        # Recipient's email from email message
        to = split_sender[-1]
        
        # We want to schedule a meeting
        if "SCHEDULE_MT" in response:

            # Recipient's name from email message
            speaker = ""

            # don't include last item (which is email) when iterating
            for speaker_subname in split_sender[:-1]:
                speaker += speaker_subname.capitalize() + " "

            # Meeting link 
            link = schedule_meeting(start_time, end_time, meeting_duration, speaker)
            subject = "Meeting Scheduled!"
            msgHtml = "Hi {speaker}!<br/>I scheduled a meeting for us using this link: {link}. See you there!".format(speaker=speaker, link=link)
            gmail_send(user_email, to, subject, msgHtml, None)

        # Small talk
        else:
            # Sending messages
            subject = "Thanks for reaching out!"
            msgHtml = response
            gmail_send(user_email, to, subject, msgHtml, None)

main()