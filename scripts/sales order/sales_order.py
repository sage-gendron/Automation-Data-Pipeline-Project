# sales_order.py
import pandas as pd
import xlwings as xw
from rename import rename
"""
author: Sage Gendron
Extract data from the engineered schedule and quote sheets in the project file. Data is simplified and transformed into 
a single sales order .csv file for order entry to directly upload into the enterprise SQL database.

Only generate_sales_order() called directly from an Excel project file by the Customer Service department.
"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
sorted_package_list: list[str] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
                                  'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',
                                  'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']


def extract_schedule_data(wb_sch):
    """
    Extracts information from schedule required to generate a sales order.

    :param xw.Book wb_sch: Book object containing a schedule sheet
    :return: engr_components - dictionary of engineered components and quantities organized by package keys
    :rtype: dict
    """
    # grab qty, pkg_key, balance_component columns from schedule and turn into pandas.DataFrame object
    df: pd.DataFrame = pd.read_excel(wb_sch.fullname, sheet_name='SCHEDULE', header=0, usecols='B,F,N', skiprows=31,
                                     nrows=1000)
    df.dropna(thresh=3, inplace=True)
    df.reset_index(drop=True)

    # transform DataFrame columns into transposed lists
    eq_qty_list: list[int] = df['qty'].values.tolist()
    pkg_key_list: list[str] = df['pkg_key'].values.tolist()
    engr_component_list: list[str] = df['engr_component'].values.tolist()

    # create dict of engineered components to be attributed by pkg key
    engr_components: dict = {}
    qty: int
    pkg: str
    engr_cmp: str
    for qty, pkg, engr_cmp in zip(eq_qty_list, pkg_key_list, engr_component_list):
        # filter out 0 quantity rows and empty rows
        if type(pkg) is float or qty == 0 or type(engr_cmp) is float:
            continue
        # add the engineered component and quantity to the dictionary
        if pkg not in engr_components.keys():
            engr_components[pkg] = {engr_cmp: qty}
        else:
            engr_components[pkg][engr_cmp] = engr_components[pkg].get(engr_cmp, 0) + qty

    return engr_components


def extract_quote_data(wb_qte):
    """
    Extracts all information from the quote spreadsheet required to generate a sales order.

    :param xw.Book wb_qte: Book object containing a generated quote sheet
    :return:
        - part_dict - dictionary mapping a list of part numbers to package keys
        - qty_dict - dictionary mapping a list of quantities to package keys
        - price_dict - dictionary mapping a list of prices to package keys
    :rtype: (dict, dict, dict)
    """
    # create pandas.DataFrame object from quote file to dynamically process packages
    df: pd.DataFrame = pd.read_excel(wb_qte.fullname, sheet_name='QUOTE', header=0, usecols='E:L', skiprows=12,
                                     nrows=620)
    df.dropna(thresh=4, inplace=True)

    # transform DataFrame columns into transposed lists
    package_quantities: list[int] = df['pkg qty'].values.tolist()
    part_numbers: list[str] = df['parts'].values.tolist()
    part_quantities: list[int] = df['qty'].values.tolist()
    part_prices: list[float] = df['net price'].values.tolist()
    package_prices: list[float] = df['pkg price'].values.tolist()

    # instantiate structures for simple look up of quote data later
    part_dict: dict[str, list[str]] = {}
    qty_dict: dict[str, list[int]] = {}
    price_dict: dict[str, list[float]] = {}

    # instantiate helper variables that only change when a quoted package has completed its loop
    current_pkg: str = ''
    current_pkg_qty: int = 0

    # loop through part numbers in quote to identify which packages should be included in SO
    pt: str
    for pkg_qty, pt, pt_qty, pt_price, pkg_price in zip(
            package_quantities, part_numbers, part_quantities, part_prices, package_prices):
        if type(pt) is not float and pt.startswith('PACK'):
            # handle zero quantity packages (ADDs, ALTs, 0 qty releases) and reset helper variables
            if pkg_qty == 0.0:
                current_pkg = ''
                current_pkg_qty = 0
                continue

            # only proceed/include if the package net is > 0
            if pkg_price > 0.0:
                # handle packages with keys > Z (ex. AA)
                current_pkg = pt[-2] if len(pt) == 10 else f"A{pt[-2]}"
                # apply package quantity to local variable for multiplication at each part within package
                current_pkg_qty = pkg_qty
                # instantiate empty lists (at dict key pkg_key) for parts, quantities, and nets for the current pkg_key
                part_dict[current_pkg]: list[str] = []
                qty_dict[current_pkg]: list[int] = []
                price_dict[current_pkg]: list[float] = []

        # for each package, where current_pkg local variable is not blank, add parts, qtys, net prices to previously
        # instantiated lists as required
        elif current_pkg != '' and type(pt) is not float and pt != '' and pt[:4] not in ('AUX1', 'AUX2'):
            part_dict[current_pkg].append(pt)
            qty_dict[current_pkg].append(pt_qty * current_pkg_qty)
            price_dict[current_pkg].append(pt_price)

    return part_dict, qty_dict, price_dict


def assign_engineered_components(engr_components, part_dict, qty_dict, price_dict):
    """
    Takes the engineered components from the engineered schedule and merges them with the quoted components in package
    key order (this ordinality is simpler for production to keep track of).

    :param engr_components: dictionary of engineered components and quantities organized by package keys
    :param part_dict: dictionary mapping a list of part numbers to package keys
    :param qty_dict: dictionary mapping a list of quantities to package keys
    :param price_dict: dictionary mapping a list of prices to package keys
    :return:
        - part_dict - dictionary mapping a list of part numbers to package keys
        - qty_dict - dictionary mapping a list of quantities to package keys
        - price_dict - dictionary mapping a list of prices to package keys
    :rtype: (dict, dict, dict)
    """
    # loop through package keys in one of the three dictionaries (which should have the same keys)
    let: str
    for let in part_dict:
        # if no engineered components found for that package key, skip
        if let not in engr_components:
            continue
        # loop through components and quantities per package to append to sales order (by package) if required
        total_qty: int = 0
        pn: str
        pn_qty: int
        for pn, pn_qty in engr_components[let].items():
            aux_type: str = ''
            if type(pn) is float or not pn.startswith('AUX'):
                continue

            # add the engineered component, its quantity, $0 and no notes to the sales order part list
            part_dict[let].append(pn)
            qty_dict[let].append(pn_qty)
            price_dict[let].append(0.0)
            # grab first two characters of the balance component for cartridge supplementary parts
            aux_type = pn[:3]
            total_qty += pn_qty

        # append extra part if AUX1 indicated
        if aux_type == 'AUX1':
            part_dict[let].append('screw001')
            qty_dict[let].append(total_qty)
            price_dict[let].append(0.0)
        # append extra parts if AUX2 indicated (1 of first auxiliary part, 2 of second)
        elif aux_type == 'AUX2':
            part_dict[let].extend(['screw002', 'washer001'])
            qty_dict[let].extend([total_qty, total_qty * 2])
            price_dict[let].extend([0.0, 0.0])

    return part_dict, qty_dict, price_dict


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
    engr_components = extract_schedule_data(wb)

    # extract required information from the quote document and transform to organized structures
    part_dict, qty_dict, price_dict = extract_quote_data(wb)

    # assign engineered components to kits if required
    part_dict, qty_dict, price_dict = assign_engineered_components(engr_components, part_dict, qty_dict, price_dict)

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

    # create equivalent length lists in raw dataframe format to be sent to excel
    outfile_df: pd.DataFrame = pd.DataFrame(list(zip(pn_list, qty_list, net_list)))
    # attempt to send to excel, else throw an error if sales order file was originally created, but still open
    try:
        outfile_df.to_excel(so_file, index=False, header=False)
    except PermissionError:
        raise Exception('Sales Order file is still open. Please close the sales order file and try again.')
