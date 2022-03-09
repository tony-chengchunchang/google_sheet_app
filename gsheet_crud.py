from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import traceback
import datetime

SPREADSHEET_ID = '1SxbT2Gubgd9GkOK4jGcUja7QjognFcjELUfy6o4TOvE'
RANGE_NAME = 'fb_err_log!A2:I'


class GsheetHandler:
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    def __init__(self, spreadsheet_id=SPREADSHEET_ID):
        self.service = self.setup()
        self.spreadsheet_id = spreadsheet_id
        
    def setup(self):
        """Shows basic usage of the Sheets API.
        Prints values from a sample spreadsheet.
        """
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)

        return service

    def get_values(self, data_range=RANGE_NAME):
        # Call the Sheets API
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.spreadsheet_id,
                                    range=data_range).execute()
        values = result.get('values', [])

        return values
    
    
    def get_batch_values(self, data_ranges):
        """
        data_ranges:List
        """
        sheet = self.service.spreadsheets()
        result = sheet.values().batchGet(spreadsheetId=self.spreadsheet_id, ranges=data_ranges).execute()
        values = result.get('valueRanges', [])
        
        return values

    
    def update_values(self, values, range_=RANGE_NAME, majorDimension='COLUMNS'):
        """
        range_: str, A1 notation
        values: List of List, corresponding to majorDimension
        majorDimension: 'COLUMNS' or 'ROWS'
        """
        body = {
            "range": range_,
            'majorDimension':majorDimension,
            "values": values
        }

        sheet = self.service.spreadsheets()
        res = sheet.values().update(
            spreadsheetId=self.spreadsheet_id, 
            range=range_,
            valueInputOption='RAW',
            body=body
        ).execute()

        return res
    
    def clear_sheet(self, range_=RANGE_NAME):
        sheet = self.service.spreadsheets()
        res = sheet.values().clear(
            spreadsheetId=self.spreadsheet_id, 
            range=range_,
        ).execute()
        
        return res
        
        
        
        