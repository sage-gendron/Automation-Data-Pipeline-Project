# sales_order.py
import pandas as pd
import xlwings as xw
from rename import rename
"""
author: Sage Gendron
Extract data from the engineered schedule and quote to be transformed into a single sales order spreadsheet to allow
order entry to directly upload the file into the enterprise SQL database.
"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
sorted_package_list: list[str] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
                                  'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',
                                  'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']


def extract_schedule_data(wb_sch):
    """
    www

    :param xlwings.Book wb_sch:
    :return: engr_components -
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
        if type(pkg) is float or qty == 0 or type(engr_cmp) is float:
            continue
        if pkg not in engr_components.keys():
            engr_components[pkg] = {engr_cmp: qty}
        else:
            engr_components[pkg][engr_cmp] = engr_components[pkg].get(engr_cmp, 0) + qty

    return engr_components


def extract_quote_data(wb_qte):
    """
    Extracts all information from the quote spreadsheet required to generate a sales order.

    :param xlwings.Book wb_qte:
    :return:
        - package_quantities - list of package quantities extracted from quote spreadsheet
        - part_numbers - list of part numbers extracted from quote spreadsheet (includes break lines between packages)
        - part_quantities - list of individual part quantities extracted from quote spreadsheet
        - part_prices - list of individual part prices extracted from quote spreadsheet
        - package_prices - list of package prices extracted from quote spreadsheet
    :rtype: (list, list, list, list, list)
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

    return package_quantities, part_numbers, part_quantities, part_prices, package_prices


def generate_sales_order():
    """
    Takes Excel version of quote file, extracts part numbers, part quantities, kit quantities, and net prices, and
    creates a new Excel file with only that information multiplied accordingly.
    The new Excel file is saved in the calling folder and is ready for SQL import.

    :return: None - Saves new file in active folder.
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()

    #
    engr_components = extract_schedule_data(wb)

    #
    package_quantities, part_numbers, part_quantities, part_prices, package_prices = extract_quote_data(wb)

    #
    quote_pkg_pndict: dict = {}
    quote_pn_qtydict: dict = {}
    quote_pn_netdict: dict = {}

    #
    index: int = 0
    current_pkg: str = ''
    current_pkg_qty: int = 0

    # loop through part numbers in quote to identify which packages should be included in SO
    pt: str
    for pkg_qty, pt, pt_qty, pt_price, pkg_price in zip(
            package_quantities, part_numbers, part_quantities, part_prices, package_prices):
        if type(pt) is not float and pt.startswith('PACK'):
            # handle zero quantity packages (ADDs, ALTs, 0 qty releases)
            if package_quantities[index] == 0.0:
                current_pkg = ''
                current_pkg_qty = 0
                index += 1
                continue

            # only proceed/include if the package net is > 0
            if package_prices[index] > 0.0:
                # handle packages with keys > Z (ex. AA)
                current_pkg = pt[-2] if len(pt) == 10 else f"A{pt[-2]}"

                # apply package quantity to local variable for multiplication at each part within package
                current_pkg_qty = package_quantities[index]

                # instantiate empty lists (at dict key pkg_key) for parts, quantities, and nets for the current pkg_key
                quote_pkg_pndict[current_pkg]: list[str] = []
                quote_pn_qtydict[current_pkg]: list[int] = []
                quote_pn_netdict[current_pkg]: list[float] = []

        # for each package, where current_pkg local variable is not blank, add parts, qtys, net prices to previously
        # instantiated lists as required
        elif current_pkg != '' and type(pt) is not float and pt != '' and pt[:2] not in ('AUX1', 'AUX2'):
            quote_pkg_pndict[current_pkg].append(pt)
            quote_pn_qtydict[current_pkg].append(part_quantities[index] * current_pkg_qty)
            quote_pn_netdict[current_pkg].append(part_prices[index])

        index += 1

    #
    let: str
    for let in quote_pkg_pndict:
        # assign balance components to kits if required
        if let not in engr_components:
            continue
        #
        i: int = 0
        for k, v in engr_components[let].items():
            aux_type = None
            if type(k) is float:
                continue
            if not k.startswith('AUX'):
                continue

            # add the engineered component, its quantity, $0 and no notes to the sales order part list
            quote_pkg_pndict[let].append(k)
            quote_pn_qtydict[let].append(v)
            quote_pn_netdict[let].append(0.0)
            # grab first two characters of the balance component for cartridge supplementary parts
            aux_type = k[:2]
            i += v

        # append extra part if AUX1 indicated
        if aux_type == 'AUX1':
            quote_pkg_pndict[let].append('screw')
            quote_pn_qtydict[let].append(i)
            quote_pn_netdict[let].append(0.0)
        # append extra parts if AUX2 indicated (1 of first auxiliary part, 2 of second)
        elif aux_type == 'AUX2':
            quote_pkg_pndict[let].extend(['screw', 'washer'])
            quote_pn_qtydict[let].extend([i, i*2])
            quote_pn_netdict[let].extend([0.0, 0.0])

    pn_list: list[str] = []
    qty_list: list[int] = []
    net_list: list[float] = []
    # remove empty strings from lists and change into simple lists in alphabetical package order for write to Excel
    pkg_key: str
    for pkg_key in sorted_package_list:
        if pkg_key not in quote_pkg_pndict:
            continue
        for pt in quote_pkg_pndict[pkg_key]:
            pn_list.append(pt)
        for qty in quote_pn_qtydict[pkg_key]:
            qty_list.append(qty)
        for net in quote_pn_netdict[pkg_key]:
            net_list.append(net)

    # generate sales order filename
    so_file: str = rename(wb, 'SALES ORDER', 'xlsx')

    # create equivalent length lists in raw dataframe format to be sent to excel
    outfile_df: pd.DataFrame = pd.DataFrame(list(zip(pn_list, qty_list, net_list)))
    # attempt to send to excel, else throw an error if sales order file was originally created, but still open
    try:
        outfile_df.to_excel(so_file, index=False, header=False)
    except PermissionError:
        raise Exception('Sales Order file is still open. Please close the sales order file and try again.')
