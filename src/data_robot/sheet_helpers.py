

# https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#GridRange
# reusable sheet range
grid_range = {
        "sheetId": 0,
        "startRowIndex": 0,
        "endRowIndex": 0,
        "startColumnIndex": 0,
        "endColumnIndex": 0
    }

def validate_sheet_structure(sheet_data=None):
    """

    :param sheet_data:
    :return:
    """
def get_sheet_header_range(sheet_id=None, sheet_column_count=None):
    """

    :param sheet_columns:
    :return:
    """
    grid_range["sheetId"] = sheet_id
    grid_range["startRowIndex"] = 0
    grid_range["endRowIndex"] = 1
    grid_range["startColumnIndex"] = 0
    grid_range["endColumnIndex"] = sheet_column_count + 1

    return grid_range

def get_sheet_dimensions(sheet_response=None):
    """
    extract sheet dimensions

    :param sheet_response: response from sheets api
    :return: rows and columns in a list
    """
    rows, columns = [0, 0]
    try:
        for prop_key, prop_value in \
                sheet_response['properties']['gridProperties'].items():
            if prop_key == 'rowCount':
                rows = prop_value
            elif prop_key == 'columnCount':
                columns = prop_value
    except Exception as error:
        print(error.args)

    return rows, columns
