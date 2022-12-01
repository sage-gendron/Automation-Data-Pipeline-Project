# scripts/salesorder/extract.py
"""
author: Sage Gendron
Functions to extract data from the project Excel file (schedule and quote) and transforms the data into a usable data
structure.
"""
import pandas as pd


def schedule_data(wb_sch):
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


def quote_data(wb_qte):
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
