# scripts/smartsheet_utils/smartsheet_dump_columns.py
"""
author: Sage Gendron
Takes a Smartsheet sheet ID and an API key and dumps the column IDs (which are not available via Smartsheet's web UI)
into a .json file to be accessed for updating Smartsheet with information from the automated Excel estimating process.
"""
import json
import os
import smartsheet

# Smartsheet API key to connect to server
sg_api_key: str = '###'

# Unique ID for estimating Smartsheet sheet
sheet_id: str = '@@@'


def dump_column_ids(sheet):
    """
    Takes a smartsheet sheet object, downloads column IDs and produces a .json file saved to the same location this file
    is in.

    :param smartsheet.Smartsheet.models.Sheet sheet: smartsheet instantiated sheet object
    :return: None - outputs a .json file with a dictionary of columns: column IDs
    """
    # read column IDs into dictionary
    new_column_map: dict[str, int] = {}
    column: smartsheet.Smartsheet.models.Column
    for column in sheet.columns:
        new_column_map[column.title] = column.id

    # dump column map into json file for permanent lookup?
    with open('column_ids.json', 'w') as outfile:
        json.dump(new_column_map, outfile, sort_keys=True, indent=4)


def generate_important_columns():
    """
    Creates a smartsheet client, gets the sheet required, and calls function to dump column IDs into a JSON file.

    :return: None
    """
    # instantiate smartsheet client
    smartsheet_client: smartsheet.Smartsheet = smartsheet.Smartsheet(os.environ['smartsheetev:'])
    smartsheet_client.errors_as_exceptions(True)

    # create Sheets instance via smartsheet client and sheet ID number provided
    sheet: smartsheet.Smartsheet.models.Sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

    # run dump column ids function to get column ID file
    dump_column_ids(sheet)


if __name__ == '__main__':
    """
    Creates a Smartsheet client and pulls out the column ids as a json file.
    The file gets saved to the location indicated in dump_column_ids(). 
    """
    generate_important_columns()
    print('Generated column ID database. Please check your designated file location.')
    quit()
