from __future__ import print_function
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# from pyad import aduser
# import pyad.adquery

#The Active Directory checking / adding portion of the code needs to be added around line 58
#user = aduser.ADUser.from_cn("myuser") #AD credentials / specifics
#queryActiveDirectory = pyad.adquery.ADQuery()

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
# SPREADSHEET_ID = 
RANGE_NAME = 'Form Responses 1!D:D'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    credsGoogle = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credsGoogle = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not credsGoogle or not credsGoogle.valid:
        if credsGoogle and credsGoogle.expired and credsGoogle.refresh_token:
            credsGoogle.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            credsGoogle = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credsGoogle, token)

    service = build('sheets', 'v4', credentials=credsGoogle)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:  # error message if no data at all is found
        print('ERROR! Something is missing!?')
    else:
        for row in values[1:]:
            # the AD check / add will go here!
            #queryActiveDirectory.execute_query(
            #attributes = ["distinguishedName", "description"],
            #where_clause = "objectClass = '*'",
            #base_dn = "OU=users, DC=domain, DC=com")
            print(row)

if __name__ == '__main__':
    main()
