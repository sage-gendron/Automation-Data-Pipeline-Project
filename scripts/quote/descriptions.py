# quote_descriptions.py
"""
author: Sage Gendron

"""
quote_kit_descrip: dict[str, str] = {
    'A': 'H15', 'B': 'H30', 'C': 'H45', 'D': 'H60', 'E': 'H75', 'F': 'H90', 'G': 'H105', 'H': 'H120', 'I': 'H135',
    'J': 'H150', 'K': 'H165', 'L': 'H180', 'M': 'H195', 'N': 'H210', 'O': 'H225', 'P': 'H240', 'Q': 'H255', 'R': 'H270',
    'S': 'H285', 'T': 'H300', 'U': 'H315', 'V': 'H330', 'W': 'H345', 'X': 'H360', 'Y': 'H375', 'Z': 'H390',
    'AA': 'H405', 'AB': 'H420', 'AC': 'H435', 'AD': 'H450', 'AE': 'H465', 'AF': 'H480', 'AG': 'H495', 'AH': 'H510',
    'AI': 'H525', 'AJ': 'H540', 'AK': 'H555', 'AL': 'H570', 'AM': 'H585', 'AN': 'H600'}


def generate_description(json_sch):
    """
    Creates description per package based on size, coil, and cv (if info supplied).

    :param list json_sch: list of dictionaries representing the excel schedule
    :return:
        - json_sch - list of dictionaries representing the excel schedule
        - quote_descr - dictionary of {sizes: per-package descriptions} both in string format
    :rtype: (list, dict)
    """
    descrips: dict[str, str] = {}
    pack_eq_types: dict[str, list[str]] = {}
    skip_eq: list[str] = []

    # iterate over json entries with iterative variable 'kit'
    kit: int
    for kit in range(len(json_sch)):
        # pull out package key to check if description has already been created
        pkg: str = json_sch[kit]['pkg_key']

        # if the pkg key is in the descrips dict, the description has already generated, just checks eq_type
        if pkg in descrips.keys():
            if pkg not in skip_eq and (temp_eq_type := json_sch[kit]['eq_type']) not in pack_eq_types[pkg]:
                pack_eq_types[pkg].append(temp_eq_type)

        # if the pkg key is not in the dict it will form a list that will be referred to in assign_descrip_to_quote
        else:
            # instantiate package-specific helper variables
            descrip: list[str] = []
            inc_coil: bool = False
            inc_cv: bool = False

            # create eq_type list in dictionary to be appended to descrips later
            if type(temp_eq_type := json_sch[kit]['eq_type']) not in (None, float):
                pack_eq_types[pkg] = [temp_eq_type]
            elif pkg not in skip_eq:
                skip_eq.append(pkg)

            # grab relevant fields from json_sch to speed up lookups
            rate: float = json_sch[kit]['rate']
            size: str = json_sch[kit]['size']
            conn_size: str = json_sch[kit]['conn_size']
            control_size_type: str = json_sch[kit]['control_size_type']
            control_type: str = json_sch[kit]['control_type']

            # check for small size kits
            is_sm: bool
            is_sm = True if size == 'SMALL' and rate <= 5 else False

            # check for PICV/CV kits to catch cv and max rate notes in description
            cv_way: str
            if type(control_type) not in (None, float):
                # if controls by HCI
                if 'HCI' in control_type:
                    # if PICV, make cv_way PICV in lieu of '2-WAY' else grab cv_way from schedule
                    cv_way = 'PICV' if 'PICV' in control_type else json_sch[kit]['cv_way']
                # if controls by other
                else:
                    # if PICV, make cv_way PICV in lieu of '2-WAY' else grab cv_way from schedule
                    cv_way = 'PICV' if 'PICV' in control_type else json_sch[kit]['cv_way']
            else:
                # default to not including any of the below 3 items
                cv_way = None

            # this is splitting the control_size_type at the space and solely grabbing the size component
            control_size: str
            try:
                cv_split: list[str] = control_size_type.split()
                control_size = cv_split[0]
            except AttributeError:
                control_size = size

            # if not stacked and coil or control valve sizes do not equal runout size, include in description
            if size != conn_size and conn_size != 'TBD':
                inc_coil = True
            if size != control_size and control_size != 'TBD':
                inc_cv = True

            # start with system size
            descrip.append(size)

            # if compact add the word before control valve type
            if is_sm:
                descrip.append('Small')

            # if not a no-control-valve kit, include 2-way, 3-way, or PICV
            if cv_way in ('2-WAY', '3-WAY', 'PICV'):
                descrip.append(cv_way)

            # add the word kits always (also always plural)
            descrip.append('Kits')

            # if coil and cv sizes both required, handle differently than if just one of the two
            if inc_coil and inc_cv:
                descrip.append(f"({conn_size} Connection, {control_size} Control)")
            elif inc_coil:
                descrip.append(f"({conn_size} Connection)")
            elif inc_cv:
                descrip.append(f"({control_size} Control)")

            # join the description with spaces (EQ types to be added later)
            descrips[pkg] = ' '.join(descrip)

    complete_descrips: dict[str, str] = {}
    kit: int
    for kit in range(len(json_sch)):
        # assign description to json_sch to be printed in data packet
        pkg: str = json_sch[kit]['pkg_key']
        if pkg not in complete_descrips.keys():
            if pkg not in skip_eq:
                # compile equipment types per package into a single string
                eq_types: str = ', '.join(pack_eq_types[pkg])
                # add equipment types to the pre-generated description for each package and compile into single string
                complete_descrips[pkg] = f"{descrips[pkg]} ({eq_types})"
            else:
                complete_descrips[pkg] = descrips[pkg]
        json_sch[kit]['quote_descrip'] = complete_descrips[pkg]

    return json_sch, complete_descrips


def assign_descrip_to_quote(wb, descrip_dict):
    """
    Iterates through description dict and assigns values to keyed packages in global variable quote_kit_descrip.
    **This function is non-dynamic and iterates through quote_kit_descrip global variable at top of file**

    :param xw.Book wb: xlwings Book representing combination schedule/quote file with sheet[1] being the quote
    :param dict descrip_dict: dict of package key: description
    :return: None - changes made to Excel file
    """
    # iterate through generated descriptions and assign package descriptions to quote sheet
    try:
        k: str
        v: str
        for k, v in descrip_dict.items():
            wb.sheets['QUOTE'].range(quote_kit_descrip[k]).value = v
    except KeyError:
        raise Exception('Please check to make sure your package keys are all upper case and between \'A\' and \'AN\'.')
