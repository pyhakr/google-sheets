import os
import sys
import httplib2
# import argparse

from oauth2client import tools
from apiclient import discovery
from oauth_helpers import get_credentials
from app_config import get_app_config
from sheet_helpers import *

# based on https://developers.google.com/sheets/api/quickstart/python

def main():
    """
    """


    app_config = get_app_config(
        file=os.path.abspath('../../config/app_config.json'))

    MOVING_AVERAGE = 'Moving Average'
    SHEET_ID = app_config['sheet_id']
    SCOPES = app_config['sheet_scopes']
    CLIENT_SECRET_FILE = os.path.abspath('../../config/{}'.format(
        app_config['client_secret_file']))
    APPLICATION_NAME = app_config['app_name']

    credentials = get_credentials(SCOPES=SCOPES,
                                  CLIENT_SECRET_FILE=CLIENT_SECRET_FILE,
                                  APPLICATION_NAME=APPLICATION_NAME)

    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    request = service.spreadsheets().get(spreadsheetId=SHEET_ID)
    response = request.execute()

    # for now assume first sheet has visitor data
    row_count, column_count = get_sheet_dimensions(response['sheets'][0])

    SHEET_HEADERS = get_sheet_header_range(0, column_count)
    request_body = {
        'data_filters': [
            {
                "gridRange": SHEET_HEADERS
            }
        ]
    }
    request = service.spreadsheets().values().batchGetByDataFilter(
        spreadsheetId=SHEET_ID,
        body=request_body)

    sheet_headers = request.execute()

    sheet_header_list = sheet_headers['valueRanges'][0]['valueRange']['values'][0]
    date_header_postion = sheet_header_list.index('Date')
    visitors_header_position = sheet_header_list.index('Visitors')
    visitor_column = chr(65 + visitors_header_position)
    moving_avg_position = None
    moving_avg_column = None
    update_body = None

    # put moving average in first empty column
    if MOVING_AVERAGE not in sheet_header_list:
        for header_pos in range(1, len(sheet_header_list) + 1):
            header_range = ''.join([chr(64 + header_pos), '1'])
            request = service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,range=header_range)
            next_header_data = request.execute()
            if 'values' not in next_header_data.keys():
                update_body = {
                    "range": header_range,
                    "values": [[MOVING_AVERAGE]]
                }
                request = service.spreadsheets().values().update(
                    spreadsheetId=SHEET_ID,
                    range=header_range,
                    valueInputOption='RAW',
                    body=update_body)
                request.execute()

                moving_avg_column = header_range[0]
                break

        # when just Dates and Visitors columns exist
        if len(sheet_header_list) == 2:
            update_body = {
                "range": "C1",
                "values": [[MOVING_AVERAGE]]
            }
            request = service.spreadsheets().values().update(
                spreadsheetId=SHEET_ID,
                range="C1",
                valueInputOption='RAW',
                body=update_body)
            request.execute()
            moving_avg_position = 3
            moving_avg_column = chr(64 + moving_avg_position)

    else:
        moving_avg_position = sheet_header_list.index('Moving Average')
        moving_avg_column = chr(65 + moving_avg_position)

    request = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range = ':'.join([visitor_column + '1', visitor_column + str(row_count)]))
    response = request.execute()

    # flatten the response values to single list
    visitor_values = [v[0] for v in response['values'][1:]]

    total_visitors = 0
    moving_averages = []
    for value in visitor_values:
        total_visitors += int(value)
        moving_averages.append(total_visitors // (visitor_values.index(value) + 1))

    batch_update_body = {
        "valueInputOption": "RAW",
        "data": [
            {
                "range": ':'.join([moving_avg_column + "2",
                                   moving_avg_column + str(len(moving_averages) + 1)]),
                "majorDimension": "COLUMNS",
                "values": [moving_averages]

            }
        ]
    }

    request = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body=batch_update_body)
    response = request.execute()
    print(response)



if __name__ == '__main__':
    main()