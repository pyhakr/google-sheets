import os
import sys
import json
import httplib2
# import argparse

from oauth2client import tools
from apiclient import discovery
from googleapiclient.errors import HttpError
from .oauth_helpers import get_credentials
from .sheet_helpers import *

# based on https://developers.google.com/sheets/api/quickstart/python

def main():
    """
    process google sheet based on sheetId in app config
    calculate moving average of visitors
    """

    # read in config from standard location
    app_config = get_app_config(
        file=os.path.abspath('../../config/app_config.json'))

    MOVING_AVERAGE = 'Moving Average'
    SHEET_ID = app_config['sheet_id']
    SCOPES = app_config['sheet_scopes']
    CLIENT_SECRET_FILE = os.path.abspath('../../config/{}'.format(
        app_config['client_secret_file']))
    APPLICATION_NAME = app_config['app_name']

    # google oauth setup
    credentials = get_credentials(SCOPES=SCOPES,
                                  CLIENT_SECRET_FILE=CLIENT_SECRET_FILE,
                                  APPLICATION_NAME=APPLICATION_NAME)

    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    # request sheet properties
    try:
        request = service.spreadsheets().get(spreadsheetId=SHEET_ID)
        response = request.execute()
    except HttpError as he:
        print(json.dumps(he.content.decode(), sort_keys=True, indent=4))
        sys.exit(1)

    # store sheet dimensions
    row_count, column_count = get_sheet_dimensions(response['sheets'][0])


    # update gridRange values for batch get request
    SHEET_HEADERS = get_sheet_header_range(0, column_count)

    # construct body for batch get request
    request_body = {
        'data_filters': [
            {
                "gridRange": SHEET_HEADERS
            }
        ]
    }

    # get the sheet headers in a list
    request = service.spreadsheets().values().batchGetByDataFilter(
        spreadsheetId=SHEET_ID,
        body=request_body)

    sheet_headers = request.execute()

    #extract header values from response
    sheet_header_list = sheet_headers['valueRanges'][0]['valueRange']['values'][0]

    # storing header positions used later on
    date_header_postion = sheet_header_list.index('Date')
    visitors_header_position = sheet_header_list.index('Visitors')
    visitor_column = chr(65 + visitors_header_position)
    moving_avg_position = None
    moving_avg_column = None
    update_body = None

    # when Moving Average column is missing
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

                moving_avg_position = header_pos
                moving_avg_column = chr(64 + header_pos)
                # moving_avg_column = header_range[0]
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
        # moving avg column exists, so just store data we'll use later on
        moving_avg_position = sheet_header_list.index('Moving Average')
        moving_avg_column = chr(65 + moving_avg_position)

    # pull all the visitor data out
    request = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range = ':'.join([visitor_column + '1', visitor_column + str(row_count)]))
    response = request.execute()


    # extract the visitor values and handle bad data
    visitor_values = []
    for visitor_value in response['values'][1:]:
        if len(visitor_value) == 1:
            try:
                visitor_value = int(visitor_value[0])
                visitor_values.append(visitor_value)
            except ValueError:
                visitor_values.append(0)
                continue
        else:
            visitor_values.append(0)

    # calculate the moving averages
    # use last known good moving average for bad data
    total_visitors = 0
    moving_average = 0
    moving_averages = []
    for value in visitor_values:
        if value == 0:
            moving_averages.append(moving_average)
            continue
        total_visitors += value
        moving_average = total_visitors // (visitor_values.index(value) + 1)
        moving_averages.append(moving_average)

    # this ensures the proper number format for the moving avg column
    batch_format_body = \
    {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 1,
                        "endRowIndex": 11,
                        "startColumnIndex": moving_avg_position - 1,
                        "endColumnIndex": moving_avg_position,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER",
                                "pattern": "######"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
        }
      ]
    }

    # format the moving avg column for numbers
    request = service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body=batch_format_body
    )
    response = request.execute()
    #print(response)

    update_range = ':'.join([moving_avg_column + "2",
                             moving_avg_column + str(len(moving_averages) + 1)])

    # construct body for batch moving average update
    batch_update_body = {
        "valueInputOption": "RAW",
        "data": [
            {
                "range": update_range,
                "majorDimension": "COLUMNS",
                "values": [moving_averages]

            }
        ],
        "includeValuesInResponse": True
    }

    # add moving averages in one shot
    request = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body=batch_update_body)
    response = request.execute()
    return response, update_range, moving_averages


if __name__ == '__main__':
    main()