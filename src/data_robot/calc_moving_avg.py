import os
import sys
import httplib2
# import argparse

from oauth2client import tools
from apiclient import discovery
from oauth_helpers import get_credentials
from app_config import get_app_config

# based on https://developers.google.com/sheets/api/quickstart/python

def main():
    """
    """

    app_config = get_app_config(
        file=os.path.abspath('../../config/app_config.json'))

    SHEET_ID=app_config['sheet_id']
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
    result = service.spreadsheets().get(spreadsheetId=SHEET_ID)
    response = result.execute()

    print(result)


if __name__ == '__main__':
    main()