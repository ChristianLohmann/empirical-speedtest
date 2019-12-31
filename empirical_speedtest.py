from __future__ import print_function

import logging
import os
import pickle
import sys

import click
import speedtest
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_RANGE_NAME = 'Sheet1!A1:E2'


def write_to_doc(results_dict, sheet_id):
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Retrieve the documents contents from the Docs service.
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    values = [
        ['reihe1', 'eins', 'zwei', 'drei', 'vier'],
        ['reihe2', 'fünf', 'sechs', 'sieben', 'acht']
    ]

    body = {
        'values': values
    }
    result = service.spreadsheets().values().update(spreadsheetId=sheet_id, range=SAMPLE_RANGE_NAME,
                                                    valueInputOption='RAW', body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))

    if not values:
        print('No data found.')
    else:
        print('Name, Major:')
    for row in values:
        # Print columns A and E, which correspond to indices 0 and 4.
        print('%s, %s' % (row[0], row[4]))


@click.command('Do a speedtest and write results to a table')
@click.option('--sheet-id', envvar='SHEET_ID', help='the id of the google spreadsheet to save results to')
def do_speedtest(sheet_id):
    servers = []
    # If you want to test against a specific server
    # servers = [1234]

    threads = None
    # If you want to use a single threaded test
    # threads = 1

    s = speedtest.Speedtest()
    s.get_servers(servers)
    s.get_best_server()
    s.download(threads=threads)
    s.upload(threads=threads)
    s.results.share()

    results_dict = s.results.dict()
    download_in_mbits = results_dict['download'] / 1024 / 1024
    upload_in_mbits = results_dict['upload'] / 1024 / 1024
    logging.info('download rate {}'.format(download_in_mbits))
    logging.info('upload rate {}'.format(upload_in_mbits))
    write_to_doc(results_dict, sheet_id)


if __name__ == '__main__':
    logging.basicConfig(stream=sys.stderr)
    logging.getLogger().setLevel(logging.INFO)

    do_speedtest()
