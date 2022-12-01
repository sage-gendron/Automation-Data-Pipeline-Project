# scripts/smartsheet_utils/create_objects.py
"""
author: Sage Gendron
Helper modules to assist in instantiating Smartsheet objects.
"""
import smartsheet

# smartsheet API keys by user
sg_api_key: str = '###'
user2_api_key: str = '###'
user3_api_key: str = '###'
user4_api_key: str = '###'

# create unique column ID dictionary for reference below
column_ids: dict[str, int] = {
    'City': 222222, 'Customer Pricing': 666000, 'Date Won': 121212, 'Engineer': 202020, 'Engineered Component': 171717,
    'Equipment Type': 191919, 'Industry': 161616, 'Phase': 272727, 'Quote Date': 101010, 'Quote Value': 777000,
    'SO #': 262626, 'Special Case 1': 181818, 'State': 232323, 'Status': 111000, 'Street Address': 212121, 'Zip': 242424
}


def get_ss_client(wb):
    """
    Uses filepath from the passed workbook to identify author and create a Smartsheet client using their API key.

    :param xw.Book wb: Excel workbook to grab the estimator's initials from
    :return: ss_c - a smartsheet client object created with the user's API key
    :rtype: smartsheet.Smartsheet
    """
    # find the file name by splitting the filepath of the workbook
    jn = wb.fullname.split('\\')[-1].split('_')
    # look for initials for quoting personnel to select the correct API key with which to instantiate the client
    ss_c: smartsheet.Smartsheet
    if 'USER2' in jn:
        ss_c = smartsheet.Smartsheet(user3_api_key)
    elif 'USER3' in jn:
        ss_c = smartsheet.Smartsheet(user2_api_key)
    elif 'USER4' in jn:
        ss_c = smartsheet.Smartsheet(user4_api_key)
    else:
        ss_c = smartsheet.Smartsheet(sg_api_key)
    ss_c.errors_as_exceptions(True)

    return ss_c


def create_cell(ss_c, col_name, wb=None, sheet=None, xl_cell=None, is_float=False, is_bool=False, cell_val=None):
    """
    Creates a new Smartsheet Cell object, passes values as required (or pulls them from a workbook), and returns
    the cell to be appended to a Row object.

    :param smartsheet.Smartsheet ss_c: Smartsheet client object
    :param str col_name: name of the column for reference from the column_ids global variable
    :param xw.Book wb: Excel workbook from which to extract data
    :param str sheet: Sheet name to reference when extracting from wb (xlwings workbook)
    :param str xl_cell: Cell reference from which to extract data within the sheet, book above
    :param bool is_float: indicates if value should be read as a float
    :param bool is_bool: indicates if value should be read as a bool
    :param cell_val: cell value, if known/is a stock value, so it can be attributed to the new cell object
    :return: new_cell - a new Cell object with value as specified
    :rtype: smartsheet.Smartsheet.models.Cell
    """
    # create the new Cell object
    new_cell: smartsheet.Smartsheet.models.Cell = ss_c.models.Cell()
    # assign the column id to the cell object, so it has a place to be updated in the row
    new_cell.column_id = column_ids[col_name]

    # assign the cells value based on if the value was passed to this function or if it needs to be pulled from excel
    if cell_val:
        new_cell.value = cell_val
    elif is_float:
        new_cell.value = float(wb.sheets[sheet.upper()].range(xl_cell).value)
    elif is_bool:
        new_cell.value = bool(wb.sheets[sheet.upper()].range(xl_cell).value)
    else:
        new_cell.value = wb.sheets[sheet.upper()].range(xl_cell).value

    # set cell strictness to false; if left to default value (True), doesn't allow server to edit datatypes and can
    # cause silent issues on row update
    new_cell.strict = False

    return new_cell
