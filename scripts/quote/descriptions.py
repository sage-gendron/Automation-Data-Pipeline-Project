# scripts/quote/descriptions.py
"""
author: Sage Gendron
Creates descriptions based on information provided on the engineered schedule. Only generates one description per
package key.
Assigns descriptions to the to-be-exported data packet and, on a different call, assigns descriptions to the quote file
by package key.
"""
# maps available description cells by package key
quote_kit_descrip: dict[str, str] = {
    'A': 'H15', 'B': 'H30', 'C': 'H45', 'D': 'H60', 'E': 'H75', 'F': 'H90', 'G': 'H105', 'H': 'H120', 'I': 'H135',
    'J': 'H150', 'K': 'H165', 'L': 'H180', 'M': 'H195', 'N': 'H210', 'O': 'H225', 'P': 'H240', 'Q': 'H255', 'R': 'H270',
    'S': 'H285', 'T': 'H300', 'U': 'H315', 'V': 'H330', 'W': 'H345', 'X': 'H360', 'Y': 'H375', 'Z': 'H390',
    'AA': 'H405', 'AB': 'H420', 'AC': 'H435', 'AD': 'H450', 'AE': 'H465', 'AF': 'H480', 'AG': 'H495', 'AH': 'H510',
    'AI': 'H525', 'AJ': 'H540', 'AK': 'H555', 'AL': 'H570', 'AM': 'H585', 'AN': 'H600'}


def generate_description(json_sch):
    """
    Creates description per package based on system size, connection, and control type/size (if info supplied).

    :param list json_sch: list of dictionaries representing the Excel engineered schedule
    :return:
        - json_sch - list of dictionaries representing the Excel engineered schedule
        - quote_descr - dictionary of {sizes: per-package descriptions} both in string format
    :rtype: (list, dict)
    """
    descriptions: dict[str, str] = {}
    pack_eq_types: dict[str, list[str]] = {}
    skip_eq: list[str] = []

    # iterate over JSON entries
    row: dict[str, ...]
    for row in json_sch:
        # pull out package key to check if description has already been created
        pkg: str = row['pkg_key']

        # if the pkg key is in descriptions, the description has already been generated, so just check eq_type
        if pkg in descriptions.keys():
            if pkg not in skip_eq and (eq_type := row['eq_type']) not in pack_eq_types[pkg]:
                pack_eq_types[pkg].append(eq_type)

        # if the pkg key is not in the dict it will form a list that will be referred to in assign_descrip_to_quote
        else:
            # instantiate package-specific helper variables
            description: list[str] = []
            inc_connection: bool = False
            inc_control: bool = False

            # create eq_type list in dictionary to be appended to descriptions later
            if type(eq_type := row['eq_type']) not in (None, float):
                pack_eq_types[pkg] = [eq_type]
            elif pkg not in skip_eq:
                skip_eq.append(pkg)

            # grab relevant fields from schedule row
            rate: float = row['rate']
            size: str = row['size']
            conn_size: str = row['conn_size']
            control_size_type: str = row['control_size_type']
            control_type: str = row['control_type']

            # check for small size kits
            is_sm: bool = True if size == 'SMALL' and rate <= 5 else False

            # check for control types kits to catch control and max rate notes in description
            control_method: str
            if type(control_type) not in (None, float):
                # if control type 2, make control_method CTRL_TYPE_2 else grab control_method from schedule
                control_method = 'CTRL_TYPE_2' if 'CTRL_TYPE_2' in control_type else row['control_method']
            else:
                # default to not including any of the below 3 items
                control_method = ''

            # this is splitting the control_size_type at the space and solely grabbing the size component
            control_size: str
            try:
                control_size = control_size_type.split()[0]
            except AttributeError:
                control_size = size

            # if connection or control sizes do not equal system size, include in description
            if size != conn_size and conn_size != 'TBD':
                inc_connection = True
            if size != control_size and control_size != 'TBD':
                inc_control = True

            # start building the description with system size
            description.append(size)

            # if small size package, add the word before control type
            if is_sm:
                description.append('Small')

            # if a package with control, include type 2, type 3, or control_type_2
            if control_method in ('TYPE-2 PKG', 'TYPE-3 PKG', 'CTRL_TYPE_2'):
                description.append(control_method)

            # add the word kits always (also always plural)
            description.append('Kits')

            # if connection and control sizes both required, handle differently than if just one of the two
            if inc_connection and inc_control:
                description.append(f"({conn_size} Connection, {control_size} Control)")
            elif inc_connection:
                description.append(f"({conn_size} Connection)")
            elif inc_control:
                description.append(f"({control_size} Control)")

            # join the description with spaces (EQ types to be added below)
            descriptions[pkg] = ' '.join(description)

    complete_descrips: dict[str, str] = {}
    row: dict[str, ...]
    for row in json_sch:
        # assign description to json_sch for export in data packet
        pkg: str = row['pkg_key']
        if pkg not in complete_descrips.keys():
            if pkg not in skip_eq:
                # compile equipment types per package into a single string
                eq_types: str = ', '.join(pack_eq_types[pkg])
                # add equipment types to the pre-generated description for each package and compile into single string
                complete_descrips[pkg] = f"{descriptions[pkg]} ({eq_types})"
            else:
                complete_descrips[pkg] = descriptions[pkg]
        row['quote_descrip'] = complete_descrips[pkg]

    return json_sch, complete_descrips


def assign_descrip_to_quote(wb, descrips_by_pkg):
    """
    Iterates through description dict and assigns values to keyed packages in global variable quote_kit_descrip.
    **This function is non-dynamic and iterates through quote_kit_descrip global variable at top of file**

    :param xw.Book wb: represents combination schedule/quote file
    :param dict descrips_by_pkg: dict of {package key: description}
    :return: None - changes made directly to Excel file
    """
    # iterate through generated descriptions and assign package descriptions to quote sheet
    try:
        k: str
        v: str
        for k, v in descrips_by_pkg.items():
            wb.sheets['QUOTE'].range(quote_kit_descrip[k]).value = v
    except KeyError:
        raise Exception('Please check to make sure your package keys are all upper case and between \'A\' and \'AN\'.')
