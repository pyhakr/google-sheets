import os
import sys
import json
import pytest
from data_robot.calc_moving_avg import main


best_case = "1bI3TlOgd3MOODLUmS8JK7GCORnJbbVgsjmSm0I6Fo7I"
bad_data = "1dpuYxITTRJfbFnnX8wy5ZSQMJ2d1Uyrbi3z0W61GJJw"
bad_data_first_last = "1FdON_ye9RKgYeWgXlB-2zJ7AVVfBSwXJLdIkhBUnqvM"
irregular_columns ="16XwJxNKBDy_MrBXJV8iQQEwQhU4-yrVxlEc9JZFEHls"
irregular_columns_bad_data = "16GtSTdJfzuD5vDepecgO5xO9Yg3YlFVOtqjVbh5k1VA"

file = os.path.abspath('../../config/app_config.json')

@pytest.fixture(scope="function", params=\
        [
            best_case,
            bad_data,
            bad_data_first_last,
            irregular_columns,
            irregular_columns_bad_data
        ])
def update_sheet_id(request):
    try:
        with open(file, 'r+') as app_config_file:
            config_data = json.load(app_config_file)
    except (json.JSONDecodeError, FileNotFoundError) as error:
        print(error.args)

    next_sheet_id = request.param
    config_data["sheet_id"] = next_sheet_id

    try:
        with open(file, 'w+') as app_config_file:
            json.dump(config_data, app_config_file)
    except (json.JSONDecodeError, FileNotFoundError) as error:
        print(error.args)

    yield next_sheet_id

def test_calc_moving_avg(update_sheet_id):
    updated_sheet_id = update_sheet_id

    response, update_range, update_values = main()

    response_sheet, response_range = response["responses"][0]["updatedRange"].split("!")
    response_values = [int(v) for v in response["responses"][0]["updatedData"]["values"][0]]

    assert updated_sheet_id == response["spreadsheetId"], "Wrong sheet updated"
    assert update_range == response_range, "Update Range invalid"
    assert update_values == response_values, "Update Values invalid"
    print(response)



