#Created by Drashti Mayurkumar Soni on 21 st august 2024

import os
import win32com.client
import requests
import re
import datetime
from openai import AzureOpenAI
import dateutil.parser
from azure.identity import DefaultAzureCredential
OPENAI_API_ENDPOINT = "https://openai-techdiv.openai.azure.com"
OPENAI_API_VERSION = "2024-02-15-preview"
OPENAI_DEPLOYMENT_NAME = "gpt-4o"  
API_KEY =
def read_latest_unread_email():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # This refers to the inbox
        messages = inbox.Items
        messages.Sort("?<ReceivedTime?>", True)  # Sort by received time, newest first
        message = messages.GetFirst()  # Get the first (latest) email
        while message.UnRead is False:
            message = messages.GetNext()  # Keep looking for unread messages
        subject = message.Subject
        body = message.Body
        sender = message.Sender
        message.UnRead = False  # Mark the email as read
        print(f"Subject: ?(subject?)?/nFrom: ?(sender?)?/nBody: ?(body?)")
        return body, sender  # Return the body of the email and the sender's email address
    except Exception as e:
        print(f"An error occurred while reading email: ?(str(e)?)")
        return None, None
# Function to send the email content as a prompt to Azure OpenAI and extract date and time
def send_to_openai(prompt):
    try:
        client = AzureOpenAI(
            azure_endpoint=OPENAI_API_ENDPOINT,
            api_key=API_KEY,
            api_version=OPENAI_API_VERSION
        )
response1 = client.chat.completions.create(
            model=OPENAI_DEPLOYMENT_NAME,
            messages=?<
                ?("role": "system", "content": "You are a helpful assistant."?),
                ?("role": "user", "content": f"Extract the date and start time from this message and minus the 6 hours and do not provide explanantion give direct answer: '?(prompt?)'"?)
            ?>,
            max_tokens=1024
        )
response2 = client.chat.completions.create(
            model=OPENAI_DEPLOYMENT_NAME,
            messages=?<
                ?("role": "system", "content": "You are a helpful assistant."?),
                ?("role": "user", "content": f"Extract the date and time from this message and minus the 6 hours and do not provide explanantion give direct answer: '?(prompt?)'"?)
            ?>,
            max_tokens=1024
        )
        # Print the full response for debugging
        date_time_info = response1.choices?<0?>.message.content
        print("Full GPT-4 Response:", date_time_info)
        # Use dateutil.parser to parse the date and time
        date_time_str = parse_date_time(date_time_info)
        if date_time_str:
            return date_time_str
        else:
            print("Could not parse the date and time.")
            return None
    except Exception as e:
        print(f"An error occurred while communicating with OpenAI: ?(str(e)?)")
        return None
# Helper function to parse date and time using dateutil.parser
def parse_date_time(text):
    try:
        # Use regular expression to search for date and time patterns
date_time_match = re.search(r'(?/d?(1,2?)?/s?/w+?/s?/d?(4?))?/s(?/d?(1,2?):?/d?(2?)?/s?<APMapm?>?(2?))', text)
        if date_time_match:
            # If a pattern is found, use dateutil.parser to parse it
date_time_str = f"?(date_time_match.group(1)?) ?(date_time_match.group(2)?)"
            parsed_date_time = dateutil.parser.parse(date_time_str)
            return parsed_date_time
        else:
            # Fallback: try to parse any date and time in the text
            return dateutil.parser.parse(text, fuzzy=True)
    except Exception as e:
        print(f"Error parsing date and time: ?(str(e)?)")
        return None
# Function to book an event in Outlook Calendar and send an invite
def book_in_calendar(date_time_info, recipient_email):
    try:
        # Add some flexibility to the start and end time (default 1-hour meeting)
        start_datetime = date_time_info
        end_datetime = start_datetime + datetime.timedelta(hours=1)
        # Create an Outlook appointment
        outlook = win32com.client.Dispatch("Outlook.Application")
        appointment = outlook.CreateItem(1)  # 1 is the outlook item type for an appointment
        appointment.Start = start_datetime
        appointment.End = end_datetime
        appointment.Subject = "Meeting Booking"
        appointment.Duration = 60  # Duration in minutes
        appointment.Location = "Your Office"  # Modify the location as needed
        # Add the recipient to the appointment
        appointment.Recipients.Add(recipient_email)
        appointment.MeetingStatus = 1  # This sets the appointment as a meeting
        appointment.Body = "Thank you for your booking. The meeting has been scheduled."
        # Send the invite
        appointment.Send()
        print(f"Meeting scheduled and invite sent to ?(recipient_email?).")
    except Exception as e:
        print(f"An error occurred while booking the calendar event: ?(str(e)?)")
if __name__ == "__main__":
    email_body, sender_email = read_latest_unread_email()
    if email_body and sender_email:
        date_time_info = send_to_openai(email_body)
        if date_time_info:
            book_in_calendar(date_time_info, sender_email.Address)
    else:
        print("No unread email found.")
