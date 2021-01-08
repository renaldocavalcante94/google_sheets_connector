from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd

class GoogleSheets():
    
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    def __init__(self):
        self._connect()

    def _connect(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.scopes)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)

        self.sheet = service.spreadsheets()


class SpreadSheet(GoogleSheets):

    def __init__(self,id):
        super().__init__()
        self.id = id

    def get_sheet(self,sheet_name,select_range=None):
        sheet_connector = self.sheet
        
        if select_range == None:
            select_range = "A:ZZ"
        
        range_name = sheet_name+"!"+select_range

        result = sheet_connector.values().get(spreadsheetId=self.id,
                                            range=range_name).execute()

        values = result.get('values', [])

        df = pd.DataFrame(values)

        columns = values[0]

        frame = []

        for row in values[1:]:
            _frame = dict()
            for column,value in zip(columns,row):
                _frame[column] = value

            frame.append(_frame)

        df = pd.DataFrame(frame)

        return df
