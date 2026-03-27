from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

def main():
    flow = InstalledAppFlow.from_client_secrets_file(
        'credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)

    gmail_service = build('gmail', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    sheet_name = "Email_Automation_Log"

    results = gmail_service.users().messages().list(
        userId='me',
        labelIds=['CATEGORY_PROMOTIONS'],
        maxResults=5
    ).execute()

    messages = results.get('messages', [])

    if not messages:
        print("No promotional emails found.")
        return

    for msg in messages:
        msg_data = gmail_service.users().messages().get(
            userId='me', id=msg['id']).execute()

        headers = msg_data['payload']['headers']

        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
        sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown Sender')
        date = next((h['value'] for h in headers if h['name'] == 'Date'), '')

        values = [[date, sender, subject, "Promotions"]]

        sheets_service.spreadsheets().values().append(
            spreadsheetId=get_sheet_id(sheet_name),
            range="Sheet1!A:D",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()

        print("Logged:", subject)

def get_sheet_id(sheet_name):
    # Replace this with your sheet ID
    return "YOUR_SHEET_ID"

if __name__ == '__main__':
    main()
