# sales_order.py
"""
author: Sage Gendron
Extract data from the engineered schedule and quote sheets in the project file. Data is simplified and transformed into
a single sales order .csv file for order entry to directly upload into the enterprise SQL database.

Only generate_sales_order() called directly from an Excel project file by the Customer Service department.
"""
import pandas as pd
import xlwings as xw

from salesorder.assign import engineered_components
from salesorder.extract import quote_data, schedule_data
from utils.rename import rename

# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
sorted_package_list: list[str] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
                                  'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',
                                  'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']


def generate_sales_order():
    """
    Takes Excel version of quote file, extracts part numbers, part quantities, kit quantities, and net prices, and
    creates a new Excel file with only that information multiplied accordingly.
    The new Excel file is saved in the calling folder and is ready for SQL import.

    :return: None - Saves new file in active folder.
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()

    # extract required information from the schedule document and transform to structured dictionary
    engr_components = schedule_data(wb)

    # extract required information from the quote document and transform to organized structures
    part_dict, qty_dict, price_dict = quote_data(wb)

    # assign engineered components to kits if required
    part_dict, qty_dict, price_dict = engineered_components(engr_components, part_dict, qty_dict, price_dict)

    # remove empty strings from lists and change into simple lists in alphabetical package order for write to Excel
    pn_list: list[str] = []
    qty_list: list[int] = []
    net_list: list[float] = []
    pkg_key: str
    for pkg_key in sorted_package_list:
        # if the package key is not present in part_dict, this key was not used, so proceed to next key
        if pkg_key not in part_dict:
            continue
        # otherwise append to lists to be pushed to a spreadsheet
        pn_list.extend(part_dict[pkg_key])
        qty_list.extend(qty_dict[pkg_key])
        net_list.extend(price_dict[pkg_key])

    # generate sales order filename
    so_file: str = rename(wb, 'SALES ORDER', 'xlsx')

    # create equivalent length lists in raw dataframe format to be sent to Excel
    outfile_df: pd.DataFrame = pd.DataFrame(list(zip(pn_list, qty_list, net_list)))
    # attempt to send to Excel, else throw an error if sales order file was originally created, but still open
    try:
        outfile_df.to_excel(so_file, index=False, header=False)
    except PermissionError:
        raise Exception('Sales Order file is still open. Please close the sales order file and try again.')
