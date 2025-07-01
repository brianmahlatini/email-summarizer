import os
import json
import re
import openai
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- CONFIG SECTION ---

# Path to your Google service account key JSON file
SERVICE_ACCOUNT_FILE = "C:/Users/l14/Downloads/strong-market-464608-i0-c0a928ed24c2.json"

# Your Google Sheet ID
SPREADSHEET_ID = "1aeYq7pS3vMRvb3FJiMmpmlOhOfsU7UUY9SB_bJ8N4do"

# OpenAI API Key loaded from environment variable for security
openai.api_key = os.getenv("OPENAI_API_KEY")
if not openai.api_key:
    raise ValueError("OpenAI API key not found. Set it in your environment variable OPENAI_API_KEY.")

# --- Google Sheets Authentication ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)
sheet = service.spreadsheets()

# --- OpenAI Call ---
def ai_process_email(email_text):
    prompt = f"""
You are a helpful assistant extracting info from customer support emails.
Extract the following from the email text:
- Summarize the email in 1-2 sentences.
- Customer name (if any).
- Urgency level (High, Medium, Low).
- Main topic or issue.

Email text:
\"\"\"{email_text}\"\"\"

Provide your answer as a JSON object with keys: summary, customer_name, urgency, topic.
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
        )
    except Exception as e:
        print(f" OpenAI API call failed: {e}")
        return None

    if not response.choices:
        print(" No choices in OpenAI response")
        return None

    text_response = response.choices[0].message.content.strip()

    # Try to parse JSON out of the response
    try:
        json_str = re.search(r"\{.*\}", text_response, re.DOTALL).group()
        data = json.loads(json_str)
        return data
    except Exception as e:
        print(" Error parsing response JSON:", e)
        print("Raw response:", text_response)
        return None

# --- Log to Google Sheets ---
def log_task_to_sheet(data):
    values = [[
        data.get('customer_name', ''),
        data.get('urgency', ''),
        data.get('topic', ''),
        data.get('summary', '')
    ]]
    body = {'values': values}
    try:
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range="Sheet1!A:D",
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()
        return result
    except Exception as e:
        print(f" Failed to log to Google Sheets: {e}")
        return None

# --- Process one email and log ---
def process_email_and_log(email_text):
    print("Processing email...")
    data = ai_process_email(email_text)
    if not data:
        print("Extraction failed.")
        return False

    print(" Extracted data:")
    print(data)

    print("Logging to Google Sheets...")
    result = log_task_to_sheet(data)
    if result:
        print("Successfully logged to sheet.")
        return True
    else:
        print(" Logging failed.")
        return False

# --- Main function for multiple emails ---
def main(emails):
    for idx, email in enumerate(emails, 1):
        print(f"\n--- Processing email #{idx} ---")
        process_email_and_log(email)

if __name__ == "__main__":
    # Example emails list — replace or expand this as needed
    test_emails = [
        """
        Hello,

        My name is Jane Doe. I'm having a serious issue with my account login, and I can't access it since yesterday.
        Please help as soon as possible! This is urgent.

        Thank you,
        Jane
        """,
        # Add more email texts here to process multiple
    ]

    main(test_emails)
