# scripts/quote/controls.py
"""
author:Sage Gendron
Handles quoting both control type 1 and 2 component types. Quoting these two control types are quite different in
practice, but much of the complexity was reduced/removed to reduce specificity.
"""


def control_type_1(pn, control_size, signal):
    """
    Uses part number, control_size, signal parameters from schedule to use control part number to create list with
    secondary control part. Secondary control parts only selected based on data from signal and control_size.

    :param str pn: control part number taken from schedule
    :param str control_size: control size taken from schedule
    :param str signal: signal type taken from schedule
    :returns:
        - control_parts - primary control part in index 0 and secondary control part in index 1 if separate
        - control_qtys - contains an int(1) for length of control_parts
    :rtype: (list, list)
    """
    # instantiate variables to be returned
    control_parts: list = []
    control_qtys: list = []

    # retrieve the second control part
    pn = control_1_parts(pn, control_size, signal)

    # add parts and part quantities to package list
    if type(pn) == str:
        control_parts.append(pn)
        control_qtys.append(1)
    else:
        control_parts.extend(pn)
        for n in range(len(pn)):
            control_qtys.append(1)

    return control_parts, control_qtys


def control_1_parts(pn, control_size, signal):
    """
    Takes given control information and selects the second control part.

    :param str pn: control part number taken from schedule
    :param str control_size: control size taken from schedule
    :param str signal: signal type taken from schedule
    :return: control_parts - control part number in index 0 and secondary control part in index 1 if separate
    :rtype: list
    """
    # instantiate secondary control part dictionaries based on control model
    ctrl_1_sm_part: dict[str, str] = {'24V': 'R-24', '120V': 'R-120'}
    ctrl_1_lg_part: dict[str, dict[str, str]] = {
        '24V': {'1': 'Y-1-24', '2': 'Y-2-24', '3"': 'Y-3-24'},
        '120V': {'1': 'Y-1-120', '2"': 'Y-2-120', '3': 'Y-3-120'}}

    # select secondary control part based on signal and provided control pn
    if pn.startswith('R'):
        control_parts = [pn, ctrl_1_sm_part[signal]]
    elif pn.startswith('S'):
        control_parts = [pn, ctrl_1_lg_part[control_size][signal]]
    # else, raise an error if control part number is not included in the above
    else:
        raise Exception('Control part number not found. Please review selected part numbers.')

    return control_parts


def control_type_2(pn, signal):
    """
    Uses control_part and signal columns from schedule to use control part number to create a list with secondary
    control part.
    Secondary control parts only selected based on data from signal and control type 2 part number.

    :param str pn: control part number taken from schedule selection
    :param str signal: signal type taken from schedule selection
    :return:
        - control_parts - control part number in index 0 and secondary control part in index 1
        - control_qtys - contains an int(1) for length of control_parts
    :rtype: (list, list)
    """
    # use control_2_parts function to get secondary control part required
    control_parts: list[str] = control_2_parts(pn, signal)

    # create list of quantities to match part numbers
    control_qtys: list[int] = []
    for _ in range(len(pn)):
        control_qtys.append(1)

    return control_parts, control_qtys


def control_2_parts(pn, signal):
    """
    Takes given control_part_2 information and selects secondary control part.

    :param str pn: control_part_2 part number taken from schedule
    :param str signal: signal type taken from schedule
    :return: control_parts - control_part_2 part number in index 0 and secondary control part in index 1
    :rtype: list
    """
    # instantiate secondary control part dictionaries based on control model
    ctrl_2_sm_part: dict[str, str] = {'24V': 'S-24', '120V': 'S-120'}
    ctrl_2_lg_part: dict[str, str] = {'24V': 'T-24', '120V': 'T-120'}

    control_parts: list[str] = [pn]

    # select secondary control part based on signal and provided control pn
    if pn.startswith('S'):
        control_parts.append(ctrl_2_sm_part[signal])
    elif pn.startswith('T'):
        control_parts.append(ctrl_2_lg_part[signal])
    # else, raise an error if the control_type_2 part number is not included in the above dictionaries
    else:
        raise Exception('Control part number not found. Please review selected part numbers.')

    return control_parts
