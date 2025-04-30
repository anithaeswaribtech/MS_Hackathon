import os
import re
import subprocess
import chainlit as cl
import ssl
from openpyxl import load_workbook
from requests.auth import HTTPBasicAuth
import pandas as pd
import json
from PyPDF2 import PdfReader
from openai import AzureOpenAI
import win32com.client
import ctypes
from openai import OpenAI
import requests
from datetime import datetime, timedelta
from collections import Counter
import shutil

ssl._create_default_https_context = ssl._create_unverified_context
# Initialize Azure OpenAI Service client with key-based authentication
endpoint = os.getenv("ENDPOINT_URL", "https://ai-*****************.openai.azure.com/")  
deployment = os.getenv("DEPLOYMENT_NAME", "gpt-4o")  
subscription_key = os.getenv("AZURE_OPENAI_API_KEY", "*******************************************")  


client = AzureOpenAI(
    azure_endpoint=endpoint,
    api_key=subscription_key,
    api_version="2024-08-01-preview",
)

#Search chat
def search_files_by_content(drive_path, search_keywords):
    """
    Search for files containing specific keywords in their content.
    """
    matching_files = []

    # Walk through the drive and list all files
    for root, dirs, files in os.walk(drive_path):
        for file in files:
            file_path = os.path.join(root, file)
            try:
                # Handle .xlsx files
                if file.endswith(".xlsx"):
                    if search_in_xlsx(file_path, search_keywords):
                        matching_files.append(file_path)
                # Handle .xls files
                elif file.endswith(".xls"):
                    if search_in_xls(file_path, search_keywords):
                        matching_files.append(file_path)
                # Handle .pdf files
                elif file.endswith(".pdf"):
                    if search_in_pdf(file_path, search_keywords):
                        matching_files.append(file_path)
                # Handle .txt files
                elif file.endswith(".txt"):
                    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                        content = f.read()
                        if any(keyword.lower() in content.lower() for keyword in search_keywords):
                            matching_files.append(file_path)
                else:
                    print(f"Unsupported file type: {file_path}")
            except Exception as e:
                print(f"Could not read file: {file_path}. Error: {e}")

    return matching_files


def search_in_xlsx(file_path, search_keywords):
    """
    Search for keywords in an .xlsx file.
    """
    try:
        workbook = load_workbook(file_path, read_only=True)
        for sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
            for row in worksheet.iter_rows(values_only=True):
                for cell in row:
                    if cell and any(keyword.lower() in str(cell).lower() for keyword in search_keywords):
                        return True
    except Exception as e:
        print(f"Could not read .xlsx file: {file_path}. Error: {e}")
    return False


def search_in_xls(file_path, search_keywords):
    """
    Search for keywords in an .xls file (requires xlrd library).
    """
    try:
        import xlrd
        workbook = xlrd.open_workbook(file_path)
        for sheet in workbook.sheets():
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    cell_value = sheet.cell_value(row_idx, col_idx)
                    if any(keyword.lower() in str(cell_value).lower() for keyword in search_keywords):
                        return True
    except Exception as e:
        print(f"Could not read .xls file: {file_path}. Error: {e}")
    return False


def search_in_pdf(file_path, search_keywords):
    """
    Search for keywords in a .pdf file.
    """
    try:
        reader = PdfReader(file_path)
        for page in reader.pages:
            text = page.extract_text()
            if any(keyword.lower() in text.lower() for keyword in search_keywords):
                return True
    except Exception as e:
        print(f"Could not read .pdf file: {file_path}. Error: {e}")
    return False



#Open files Explorer
def open_in_explorer_multiple(files):
    """
    Open Windows Explorer and highlight all matching files in their respective directories.
    """
    try:
        # Group files by their parent directory
        directories = {}
        for file_path in files:
            parent_dir = os.path.dirname(file_path)
            if parent_dir not in directories:
                directories[parent_dir] = []
            directories[parent_dir].append(file_path)

        # Open each directory and highlight the matching files
        for parent_dir, file_paths in directories.items():
            # Join all file paths into a single command for highlighting
            for file_path in file_paths:
                subprocess.run(["explorer", "/select,", os.path.normpath(file_path)], check=True)

    except Exception as e:
        print(f"Could not open files in Explorer. Error: {e}")
 
def search_outlook_emails(search_keywords):
    """
    Search for emails in Outlook containing specific keywords in their content.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the Inbox folder
        inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

        # Get all items in the Inbox
        messages = inbox.Items

        # Filter emails containing the search keywords
        matching_emails = []
        for message in messages:
            if message.Class == 43:  # Ensure it's a MailItem
                if search_keywords.lower() in message.Body.lower() or search_keywords.lower() in message.Subject.lower():
                    matching_emails.append({
                        "Subject": message.Subject,
                        "Sender": message.SenderName,
                        "ReceivedTime": message.ReceivedTime,
                        "Body": message.Body
                    })

        # Return the matching emails
        print(f"Found {len(matching_emails)} matching emails.") 
        return matching_emails
    except Exception as e:
        print(f"An error occurred: {e}")
        return []
    
#Create Folder in Outlook and Move Mails with Subject line
def create_folder_and_move_emails(folder_name, subject_keyword):
    """
    Create a folder in Outlook and move emails with a specific subject keyword into it.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the Inbox folder
        inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

        # Check if the folder already exists
        target_folder = None
        for folder in inbox.Folders:
            if folder.Name == folder_name:
                target_folder = folder
                break

        # If the folder doesn't exist, create it
        if not target_folder:
            target_folder = inbox.Folders.Add(folder_name)
            print(f"Folder '{folder_name}' created successfully.")

        # Get all items in the Inbox
        messages = inbox.Items

        # Iterate through emails and move matching ones
        moved_count = 0
        for message in messages:
            if message.Class == 43:  # Ensure it's a MailItem
                if subject_keyword.lower() in message.Subject.lower():
                    message.Move(target_folder)
                    moved_count += 1

        print(f"Moved {moved_count} email(s) to the folder '{folder_name}'.")
    except Exception as e:
        print(f"An error occurred: {e}")

#Find Overlapping Meetings in Outlook
def find_overlapping_meetings():
    """
    Find overlapping meetings scheduled for today in the Outlook calendar.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the Calendar folder
        calendar = outlook.GetDefaultFolder(9)  # 9 refers to the Calendar folder

        # Get today's date range
        today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        today_end = today_start + timedelta(days=1)

        # Filter calendar items for today's meetings
        calendar_items = calendar.Items
        calendar_items.Sort("[Start]")
        calendar_items.IncludeRecurrences = True

        # Restrict items to today's date range
        restricted_items = calendar_items.Restrict(
            f"[Start] >= '{today_start.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{today_end.strftime('%m/%d/%Y %H:%M %p')}'"
        )

        # Convert restricted_items to a list to avoid issues with indexing
        restricted_items = list(restricted_items)

        # Check for overlapping meetings
        overlapping_meetings = []
        if len(restricted_items) < 2:
            return overlapping_meetings  # No overlapping meetings possible with fewer than 2 items

        for i in range(1, len(restricted_items)):
            prev_item = restricted_items[i - 1]
            curr_item = restricted_items[i]
            if prev_item.End > curr_item.Start:
                overlapping_meetings.append(
                    f"Meeting 1: {prev_item.Subject} ({prev_item.Start} - {prev_item.End})\n"
                    f"Meeting 2: {curr_item.Subject} ({curr_item.Start} - {curr_item.End})"
                )

        # Return overlapping meetings
        return overlapping_meetings

    except Exception as e:
        print(f"An error occurred: {e}")
        return []

def find_overlapping_meetings_tomorrow():
    """
    Find overlapping meetings scheduled for tomorrow in the Outlook calendar.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the Calendar folder
        calendar = outlook.GetDefaultFolder(9)  # 9 refers to the Calendar folder

        # Get tomorrow's date range
        tomorrow_start = (datetime.now() + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        tomorrow_end = tomorrow_start + timedelta(days=1)

        # Filter calendar items for tomorrow's meetings
        calendar_items = calendar.Items
        calendar_items.Sort("[Start]")
        calendar_items.IncludeRecurrences = True

        # Restrict items to tomorrow's date range
        restricted_items = calendar_items.Restrict(
            f"[Start] >= '{tomorrow_start.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{tomorrow_end.strftime('%m/%d/%Y %H:%M %p')}'"
        )

        # Convert restricted_items to a list to avoid issues with indexing
        restricted_items = list(restricted_items)

        # Check for overlapping meetings
        overlapping_meetings = []
        if len(restricted_items) < 2:
            return overlapping_meetings  # No overlapping meetings possible with fewer than 2 items

        for i in range(1, len(restricted_items)):
            prev_item = restricted_items[i - 1]
            curr_item = restricted_items[i]
            if prev_item.End > curr_item.Start:
                overlapping_meetings.append(
                    f"Meeting 1: {prev_item.Subject} ({prev_item.Start} - {prev_item.End})\n"
                    f"Meeting 2: {curr_item.Subject} ({curr_item.Start} - {curr_item.End})"
                )

        # Return overlapping meetings
        return overlapping_meetings

    except Exception as e:
        print(f"An error occurred: {e}")
        return []

#Locate a specific file in the system
def locate_file(file_name, search_path):
    """
    Locate a specific file in the system by searching recursively in the given directory.

    Args:
        file_name (str): The name of the file to search for.
        search_path (str): The directory path to start the search.

    Returns:
        list: A list of full paths to the located file(s).
    """
    located_files = []

    # Walk through the directory and its subdirectories
    for root, dirs, files in os.walk(search_path):
        if file_name in files:
            located_files.append(os.path.join(root, file_name))

    return located_files

# Find Most sent Receipient in Outlook
def find_most_sent_recipient():
    """
    Find the most frequently emailed recipient in the Outlook Sent Items folder.

    Returns:
        str: The email address of the most frequently emailed recipient.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the Sent Items folder
        sent_items = outlook.GetDefaultFolder(5)  # 5 refers to the Sent Items folder

        # Get all items in the Sent Items folder
        messages = sent_items.Items

        # Collect recipient email addresses
        recipients = []
        for message in messages:
            if message.Class == 43:  # Ensure it's a MailItem
                for recipient in message.Recipients:
                    recipients.append(recipient.Address)

        # Count the frequency of each recipient
        recipient_counts = Counter(recipients)

       # Find the most frequently emailed recipient
        most_sent_recipient = recipient_counts.most_common(1)
        if most_sent_recipient:
         match = re.search(r"cn=.*?-(\w+)$", most_sent_recipient[0][0])
         receipient=match.group(1) if match else most_sent_recipient[0][0]
         return f"The Most sent Recipient is: {receipient} with {most_sent_recipient[0][1]} emails."
        else:
         return "No recipients found in the Sent Items folder."

    except Exception as e:
        print(f"An error occurred: {e}")
        return "An error occurred while processing the Sent Items folder."
    
def count_emails_from_sender(sender_email):
   
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the Inbox folder
        inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

        # Get all items in the Inbox
        messages = inbox.Items

        # Filter emails from the specific sender
        count = 0
        for message in messages:
            if message.Class == 43:  # Ensure it's a MailItem
                if message.SenderEmailAddress.lower() == sender_email.lower():
                    count += 1

        return count

    except Exception as e:
        print(f"An error occurred: {e}")
        return 0
    
#Group files into personal, official, and secure categories based on keywords
def group_files(base_path):
    """
    Group files into Personal, Official, and Secure categories based on keywords.
    Files are moved into their respective folders.

    Args:
        base_path (str): The base directory to search for files.
    """
    personal_keywords = ["Anitha", "photo", "siddharth", "samyuktha", "tuition", "images"]
    official_keywords = ["report", "presentation", "meeting", "project", "work", "office", "Test Procedure"]
    secure_keywords = ["confidential", "secret", "private", "restricted", "sensitive", "bank", "password", "passwords"]

    # Define directories for grouping
    personal_dir = os.path.join(base_path, "Personal_Files")
    official_dir = os.path.join(base_path, "Official_Files")
    secure_dir = os.path.join(base_path, "Secure_Files")

    # Create directories if they don't exist
    os.makedirs(personal_dir, exist_ok=True)
    os.makedirs(official_dir, exist_ok=True)
    os.makedirs(secure_dir, exist_ok=True)

    # Walk through the base directory
    for root, dirs, files in os.walk(base_path):
        for file in files:
            file_path = os.path.join(root, file)
            try:
                # Skip directories created for grouping
                if root in [personal_dir, official_dir, secure_dir]:
                    continue

                # Check file content for keywords
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read().lower()

                    # Classify files based on keywords and move them
                    if any(keyword.lower() in content for keyword in personal_keywords):
                        shutil.move(file_path, os.path.join(personal_dir, file))
                    elif any(keyword.lower() in content for keyword in official_keywords):
                        shutil.move(file_path, os.path.join(official_dir, file))
                    elif any(keyword.lower() in content for keyword in secure_keywords):
                        shutil.move(file_path, os.path.join(secure_dir, file))

            except Exception as e:
                print(f"Could not process file: {file_path}. Error: {e}")
#Forward Email
def forward_email(email, recipient_email):
    """
    Forward an email to a specified recipient.

    Args:
        email (dict): The email to forward (contains Subject, Sender, Body, etc.).
        recipient_email (str): The recipient's email address.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # Create a new mail item

        # Set email details
        mail.To = recipient_email
        mail.Subject = f"Fwd: {email['Subject']}"
        mail.Body = f"Forwarded email:\n\n{email['Body']}"
        mail.Send()

        print(f"Email forwarded to {recipient_email}.")
    except Exception as e:
        print(f"An error occurred while forwarding the email: {e}")



def send_email_with_attachment(recipient_email, subject, body, attachment_path):
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # Create a new mail item

        # Set email details
        mail.To = recipient_email
        mail.Subject = subject
        mail.Body = body

        # Attach the file
        if os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)
        else:
            print(f"Attachment not found: {attachment_path}")
            return

        # Send the email
        mail.Send()
        print("Email sent successfully!")

    except Exception as e:
        print(f"An error occurred: {e}")

def open_file(file_path):
    """
    Open a file using the default application associated with its file type.

    Args:
        file_path (str): The full path to the file to open.
    """
    try:
        if os.path.exists(file_path):
            # Open the file using the default application
            os.startfile(file_path)  # For Windows
            print(f"File opened successfully: {file_path}")
        else:
            print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred while opening the file: {e}")

tools = [
    {
        "type": "function",
        "function": {
            "name": "search_files_by_content",
            "description": "Search File Contents",
            "parameters": {
                "type": "object",
                "properties": {
                    "drive_path": {"type": "string", "description": "drive path"},
                    "search_keywords": {"type": "string", "description": "search keywords"}
                },
                "required": ["drive_path", "search_keywords"]
            }
        }
    },
    {
        "type":"function",
        "function": {
            "name": "search_outlook_emails",
            "description": "Search Mail Contents",
            "parameters": {
                "type": "object",
                "properties": {
                   
                    "search_keywords": {"type": "string", "description": "search keywords"}
                },
                "required": ["search_keywords"]
            }
    }
    },
    {
        "type":"function",
        "function": {
            "name": "create_folder_and_move_emails",
            "description": "Create Folder and Move Mail Contents",
            "parameters": {
                "type": "object",
                "properties": {
                    "folder_name": {"type": "string", "description": "folder name"},
                    "subject_keyword": {"type": "string", "description": "subject keyword"}
                },
                "required": ["folder_name","subject_keyword"]
            }
    }
    },
    {
        "type":"function",
        "function": {
            "name": "find_overlapping_meetings",
            "description": "Find Overlapping Meetings Today"
            
    }
    },
    {
        "type":"function",
        "function": {
            "name": "find_overlapping_meetings_tomorrow",
            "description": "Find Overlapping Meetings Tomorrow"
            
    }
    },{
        "type":"function",
        "function": {
            "name": "locate_file",
            "description": "Locate a Specific file in the system",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_name": {"type": "string", "description": "file name"},
                    "search_path": {"type": "string", "description": "search path"}
                },
                "required": ["file_name","search_path"]
            }
            
    }
    },
    {
        "type":"function",
        "function": {
            "name": "find_most_sent_recipient",
            "description": "Find Most sent Receipient in Outlook"
    }
    },
    {
        "type":"function",
        "function": {
            "name": "count_emails_from_sender",
            "description": "Count how many Emails came from specific Person",
            "parameters": {
                "type": "object",
                "properties": {
                    "sender_email": {"type": "string", "description": "sender name"}
                },
                "required": ["sender_email"]
            }
            
    }
    },
    {
        "type":"function",
        "function": {
            "name": "group_files",
            "description": "Categorize the files into Personal, Official, and Secure based on keywords",
            "parameters": {
                "type": "object",
                "properties": {
                    "base_path": {"type": "string", "description": "folder path"}
                                   },
                "required": ["base_path"]
            }
            
    }
    }
]
Systemprompt = """
You are an AI-powered File Organizer.
Your primary task is to organise files and read the contents through it.
You must process drive name in **uppercase** and keywords in **uppercase" correctly identify methods.

**Instructions:**
1. If a user requests **find file with contents** , call `search_files_by_content`.
   - get all file contents in the drive and check keywords.
   - Example:
     - Query: "Find files with contents 'API KEY' in folder C:\\API"
     - drivepath : C:\\API, search_keywords : API_KEY

2. If a user requests **find mail with contents** , call `search_outlook_emails`.
   - get all mail contents in the outlook and check keywords.
   - Example:
     - Query: "Find mails with contents 'API KEY'
     - search_keywords : API_KEY
3. If a user requests **create folder with name** , call `create_folder_and_move_emails`.
   - get all mail with subject line in the outlook and move to a folder.
   - Example:
     - Query: "Create Folder with Name 'Demo' and Move Mails with Subject 'API KEY'"
     - folder_name : Demo, subject_keyword : API_KEY
4. If a user requests **find overlapping meetings today** , call `find_overlapping_meetings`.
   - get Overlapping meetings in the outlook calendar.
   - Example:
     - Query: "Find Overlapping meetings today"
     - No parameters required.
5. If a user requests **find overlapping meetings tomorrow** , call `find_overlapping_meetings_tomorrow`.
   - get Overlapping meetings in the outlook calendar.
   - Example:
     - Query: "Find Overlapping meetings tomorrow"
     - No parameters required.
6. If a user requests **locate a specific file** , call `locate_file`.
   - get specific file in the system by searching recursively in the given directory.
   - Example:
     - Query: "Find file 'Anitha.pdf' in folder C:\\"
     - file_name : Anitha.pdf, search_path : C:\\
7. If a user requests **find most sent receipient** , call `find_most_sent_recipient`.
   - get most sent receipient in the outlook sent items folder.
   - Example:
     - Query: "Find Most sent Receipient in Outlook"
     - No parameters required.
8. If a user requests **how many email from a specific person** , call `count_emails_from_sender`.
   - get number of emails received from specific sender.
   - Example:
     - Query: "How Many Emails I received from 'Anitha.Eswari@honeywell.com'"
     - sender_email : Anitha.Eswari@honeywell.com
9. If a user requests **find mail with contents** , call `search_outlook_emails`.
   - get all mail contents in the outlook and check keywords.
   - Example:
     - Query: "Find mail with contents 'API KEY'
     - search_keywords : API_KEY
10. If a user requests **create folder with name** , call `create_folder_and_move_emails`.
   - get all mail with subject line in the outlook and move to a folder.
   - Example:
     - Query: "Move Mails with Subject 'API KEY' to 'Demo' folder"
     - folder_name : Demo, subject_keyword : API_KEY
11. If a user requests **locate a specific file** , call `locate_file`.
   - get specific file in the system by searching recursively in the given directory.
   - Example:
     - Query: "Find file 'Anitha.pdf'"
     - file_name : Anitha.pdf, search_path : C:\\
12. If a user requests **find most sent receipient** , call `find_most_sent_recipient`.
   - get most sent receipient in the outlook sent items folder.
   - Example:
     - Query: "Find My Favorite Recipient in Outlook"
     - No parameters required.
13. If a user requests **categorize files based on personal, official and secure keywords** , call `group_files`.
   - categorize the files into Personal, Official, and Secure based on keywords.
   - Example:
     - Query: "Group files in folder 'C:\\GroupFiles'"
     - base_path : C:\\GroupFiles


    
"""

@cl.on_message
async def main(usermessage: str):
    # Prepare the chat prompt
    chat_prompt = [
        {
            "role": "system",
            "content": [
                {
                    "type": "text",
                    "text": Systemprompt
                }
            ]
        },
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": usermessage.content
                }
            ]
        }
    ]
    # Generate the completion
    completion = client.chat.completions.create(
        model=deployment,
        messages=chat_prompt,
        tools = tools,
        tool_choice = "auto",
        max_tokens=800,
        temperature=0.7
    )

    responses = []

    if completion.choices[0].message.tool_calls:
        for tool_call in completion.choices[0].message.tool_calls:
            function_name = tool_call.function.name
            arguments = json.loads(tool_call.function.arguments)
            if function_name == "search_files_by_content":
                result = search_files_by_content(**arguments)
                responses.append(f"**Search Performed successfully.**\n\n"+"\n".join(result))
                await cl.Message(content="\n".join(responses)).send()
                user_response_file=await cl.AskUserMessage(content="Do you want to open the file?").send()
                if user_response_file["output"].lower()== "yes":
                    open_file(result[0])  # Open the first file as an example
                    await cl.Message(content=f"File Opened Successfully").send()
          

            elif function_name == "search_outlook_emails":
             result = search_outlook_emails(**arguments)
             if result:
                print(f"Emails found: {len(result)}")
                # Format the email results into readable strings
                formatted_emails = []
                for email in result:
                    formatted_emails.append(
                        f"Subject: {email['Subject']}\n"
                        f"Sender: {email['Sender']}\n"
                        f"Received: {email['ReceivedTime']}\n"
                        f"Body: {email['Body'][:100]}..."  # Show the first 100 characters of the body
                    )
                print(formatted_emails)
                responses.append(f"**Search Performed Successfully.**\n\n" + "\n\n".join(formatted_emails))
                await cl.Message(content="\n".join(responses)).send()
                #Ask user message
                # Ask if the user wants to forward the email
                user_response = await cl.AskUserMessage(content="Do you want to forward any of these emails to someone? Please reply with 'yes' or 'no'.").send()
                print(user_response)
                if user_response["output"].lower()== "yes":
                  # Ask for the recipient's email address
                 recipient_email_response = await cl.AskUserMessage(content="Please provide the recipient's email address.").send()
                 recipient_email = recipient_email_response["output"]
                # Forward the email
                 forward_email(result[0], recipient_email)  # Forward the first email as an example
                 await cl.Message(content=f"**Email forwarded successfully to {recipient_email}.**").send()
                # Ask if the user wants to forward the email
                 user_response_1 = await cl.AskUserMessage(content="Do you want to move these emails to a folder? Please reply with 'yes' or 'no'.").send()
                 print(user_response_1)
                 if user_response_1["output"].lower()== "yes":
                   # Ask for the recipient's email address
                   folder_name_response = await cl.AskUserMessage(content="Please provide the folder name").send()
                   folder_name= folder_name_response["output"]
                   # Move the email to the specified folder
                   create_folder_and_move_emails(folder_name, result[0]['Subject'])
                   await cl.Message(content=f"Folder Successfully Created and Mails Moved to it.").send()
                   
          
            elif function_name == "create_folder_and_move_emails":
             create_folder_and_move_emails(**arguments)
             responses.append(f"**Mails Moved Successfully to the folder '{arguments['folder_name']}'.**")

            elif function_name == "find_overlapping_meetings":
                result=find_overlapping_meetings()
                responses.append("**Overlapping meetings found successfully.**\n\n"+"\n".join(result))

            elif function_name == "find_overlapping_meetings_tomorrow":
                result=find_overlapping_meetings_tomorrow()

                responses.append("**Overlapping meetings found successfully.**\n\n"+"\n".join(result))
            elif function_name == "locate_file":
                result=locate_file(**arguments)
                responses.append("**file found successfully.**\n\n"+"\n".join(result))
               
            elif function_name == "find_most_sent_recipient":
                result=find_most_sent_recipient()
                responses.append(f"**{result}")

            elif function_name == "count_emails_from_sender":
                result=count_emails_from_sender(**arguments)
                responses.append(f"**Number of Emails You have Received from the User is **{result}")

            elif function_name == "group_files":
                result=group_files(**arguments)
                responses.append(f"**Files Grouped Successfully..")
            else:
                responses.append("No matching emails were found.")
            await cl.Message(content="\n".join(responses)).send()
                
           
       
    else:
        assistant_response = completion.choices[0].message.content
        await cl.Message(content=assistant_response).send()



if __name__ == "__main__":
    from chainlit.cli import run_chainlit
    run_chainlit(__file__)