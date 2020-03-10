import pickle
import os.path
import sys
from operator import itemgetter
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# We both want to read and write to the spreadsheet
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_ID = '1XgJCDX78jg_ib_wBJUsKxRkUxmgbQ6_33l44BfRHJo0'
INPUTOPTION = 'USER_ENTERED'
RANGES = ['A4:H']
RANGE = 'Tema!A4:H'


def get_unique_el(elements, sort_by=(0, 1)):
    """Returns a list of unique list, sorted by the
    n'th element of the lists. sort_by supports tuples and ints.
    """
    uniques = [list(x) for x in set(tuple(x) for x in elements)]
    sort = sorted(uniques, key=itemgetter(*sort_by))
    return sort


def auth():
    """Perform authentication if needed.
    Ripped straight from quickstart.py from Google.
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    return sheet


def propagate_down(values, columnID):
    # Propagates the value of a given cell down through empty cells, in a given column
    for i in range(len(values)-1):
        if len(values[i+1][columnID]) == 0:
            values[i+1][columnID] = values[i][columnID]
    return values


def prepare_sheet_results(values, cols=[0, 1, 5, 6, 7]):
    # Values are returned as a list of lists, each of which contain the cell contents.
    # Exclude empty rows - they correspond to empty lists.
    values = [x for x in values if len(x) != 0]
    # propagate down the category value
    values = propagate_down(values, 0)
    # Keep only certain columns.
    usable_values = [[row[i] for i in cols] for row in values]
    return usable_values


def main():
    sheet = auth()
    result = sheet.values().get(spreadsheetId=SHEET_ID,
                                range=RANGE).execute()
    values = result.get('values', [])
    usable_values = prepare_sheet_results(values)
    print (usable_values)


if __name__ == '__main__':
    main()
