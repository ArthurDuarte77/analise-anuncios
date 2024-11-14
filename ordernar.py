
import pandas as pd
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # Read/write access
SPREADSHEET_ID = "1vMwTxe1dkYZpp9ti1C8chyd5qY3bkff3jb9PqA2LfJg"
DATA_RANGE = "DADOS!A1:I" 

def get_sheets_credentials():
    """Handles authentication with the Google Sheets API."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def update_sheet_data(spreadsheet_id, range_name, values):
    """Updates data in a specified range of a Google Sheet."""
    creds = get_sheets_credentials()
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    try:
        body = {"values": values}
        result = (
            sheet.values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption="USER_ENTERED",  # Important for direct data entry
                body=body,
            )
            .execute()
        )
        print(f"{result.get('updatedCells')} cells updated.")
    except HttpError as error:
        print(f"An error occurred while updating data: {error}")

# Path to your Excel file
file_path = os.path.join('', 'planilha_final.xlsx')  # Adjust path if necessary

# --- Modified main execution block ---
if __name__ == "__main__":

    # 1. Read data from Excel file
    try:
        df = pd.read_excel(file_path)
        # Clean your DataFrame
        df = df.fillna('')  #Fill missing values with empty strings
        for col in df.columns:
            if df[col].dtype == 'object': # only clean string columns
                df[col] = df[col].astype(str).str.replace("'", "''", regex=False)
    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        exit()  # Exit if the file doesn't exist

    # 2. Prepare data for Google Sheets update
    data_to_update = []
    header = df.columns.tolist()  # Include the header row
    data_to_update.append(header)

    for index, row in df.iterrows():
        row_data = [row[col] for col in header]
        data_to_update.append(row_data)  #Add each row from the dataframe


    # 3. Update the Google Sheet
    try:
        update_sheet_data(SPREADSHEET_ID, DATA_RANGE, data_to_update)
        print("Data updated successfully.")

    except Exception as e:
        print(f"An error occurred during update: {e}")


