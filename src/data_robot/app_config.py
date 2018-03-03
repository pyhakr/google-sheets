import json


def get_app_config(file=None):
    try:
        with open(file, 'r') as app_config_file:
            return json.load(app_config_file)
    except (json.JSONDecodeError, FileNotFoundError) as error:
        print(error.args)