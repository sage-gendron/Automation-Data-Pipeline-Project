# scripts/submittal/spec.py
"""
author: Sage Gendron
Module to assist in determining specification sheets required to be included in submittal document generation.
"""
# GLOBAL VARIABLES FOR controls()
ctrl_1_sm_part: dict[str, str] = {'24V': 'R-24.pdf', '120V': 'R-120.pdf'}
ctrl_1_lg_part: dict[str, dict[str, str]] = {
    '24V': {'1': 'Y-1-24.pdf', '2': 'Y-2-24.pdf', '3"': 'Y-3-24.pdf'},
    '120V': {'1': 'Y-1-120.pdf', '2"': 'Y-2-120.pdf', '3': 'Y-3-120.pdf'}}
ctrl_2_sm_part: dict[str, str] = {'24V': 'S-24.pdf', '120V': 'S-120.pdf'}
ctrl_2_lg_part: dict[str, str] = {'24V': 'T-24.pdf', '120V': 'T-120.pdf'}


def lg_spec(dwg_parts, spec_list_lg):
    """
    Parses through large drawings to identify spec pages required. Always includes a part 'B' spec sheet.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param list spec_list_lg: list of spec sheets identified so far
    :return: spec_list_lg - list of spec sheets including those identified for this special case 1 type
    :rtype: list
    """
    a: str = 'A-LG.pdf'
    b: str = 'B-LG.pdf'
    c: str = 'C-LG.pdf'
    d: str = 'D-LG.pdf'
    e: str = 'E-LG.pdf'
    f: str = 'F-LG.pdf'
    # loop through first portion of drawing code (ie L2ABCD) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'A' and a not in spec_list_lg:
            spec_list_lg.append(a)
        elif char == 'B' and b not in spec_list_lg:
            spec_list_lg.append(b)
        elif char == 'C' and c not in spec_list_lg:
            spec_list_lg.append(c)
        elif char == 'D' and d not in spec_list_lg:
            spec_list_lg.append(d)
        elif char == 'E' and e not in spec_list_lg:
            spec_list_lg.append(e)
    # check to ensure 'B' spec included if not already
    if b not in spec_list_lg:
        spec_list_lg.append(b)
    # check for aux spec 'F'
    if dwg_parts[1][1] == 'F' and f not in spec_list_lg:
        spec_list_lg.append(f)
    return spec_list_lg


def typ_spec(dwg_parts, is_sm, spec_list):
    """
    Loops through characters in the first half of the smart package code to come up with a list of spec sheets required.

    :param list dwg_parts: smart packaage code (drawing name) split by hyphens
    :param bool is_sm: is this a small package?
    :param list spec_list: list of spec sheets identified so far
    :return: spec_list - list of spec sheets including those identified for this special case 1 type
    :rtype: list
    """
    a: str = 'A.pdf'
    a_sm: str = 'A-SM.pdf'
    b: str = 'B.pdf'
    c: str = 'C.pdf'
    d: str = 'D.pdf'
    d_sm: str = 'D-SM.pdf'
    e: str = 'E.pdf'
    f: str = 'F.pdf'
    # loop through first portion of drawing code (ie 2ABCD) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'A':
            # check for small size
            if is_sm and a_sm not in spec_list:
                spec_list.append(a_sm)
            elif not is_sm and a not in spec_list:
                spec_list.append(a)
        elif char == 'B' and b not in spec_list:
            spec_list.append(b)
        elif char == 'C' and c not in spec_list:
            spec_list.append(c)
        elif char == 'D':
            # check for small size
            if is_sm and d_sm not in spec_list:
                spec_list.append(d_sm)
            elif not is_sm and d not in spec_list:
                spec_list.append(d)
        elif char == 'E' and e not in spec_list:
            spec_list.append(e)
    # check for aux spec 'F'
    if dwg_parts[1][2] == 'F' and f not in spec_list:
        spec_list.append(f)

    return spec_list


def sp_case_1(dwg_parts, spec_list):
    """
    Grab spec sheet names (literal strings) as required based on components in first half of the smart packaage code.
    Specifically called if the row is flagged special case 1 in the schedule.

    :param list dwg_parts: smart package code (drawing name) split by hyphens
    :param list spec_list: list of spec sheets identified so far
    :return: spec_list - list of spec sheets including those identified for this special case 1 type
    :rtype: list
    """
    a: str = 'A-SP1.pdf'
    b: str = 'B-SP1.pdf'
    c: str = 'C-SP1.pdf'
    d: str = 'D-SP1.pdf'
    e: str = 'E-SP1.pdf'
    f: str = 'F-SP1.pdf'
    # loop through first portion of drawing code (ie 2ABCD) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'A' and a not in spec_list:
            spec_list.append(a)
        elif char == 'B' and b not in spec_list:
            spec_list.append(b)
        elif char == 'C' and c not in spec_list:
            spec_list.append(c)
        elif char == 'D' and d not in spec_list:
            spec_list.append(d)
        elif char == 'E' and e not in spec_list:
            spec_list.append(e)
    # check for aux spec 'F'
    if dwg_parts[1][2] == 'F' and f not in spec_list:
        spec_list.append(f)
    return spec_list


def sp_case_2(dwg_parts, spec_list):
    """
    Loops through characters in the first half of the smart package code to come up with a list of spec sheets required.
    Specifically called if special case 2 is one of the connections to the package.

    :param list dwg_parts: smart package code (drawing name) split by hyphens
    :param list spec_list: list of spec sheets identified so far
    :return: spec_list - list of spec sheets including those identified for this special case 2 type
    :rtype: list
    """
    a: str = 'A-SP2.pdf'
    b: str = 'B-SP2.pdf'
    c: str = 'C-SP2.pdf'
    d: str = 'D-SP2.pdf'
    e: str = 'E-SP2.pdf'
    f: str = 'F-SP2.pdf'
    # loop through first portion of drawing code (ie 2ABCD) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'A' and a not in spec_list:
            spec_list.append(a)
        elif char == 'B' and b not in spec_list:
            spec_list.append(b)
        elif char == 'C' and c not in spec_list:
            spec_list.append(c)
        elif char == 'D' and d not in spec_list:
            spec_list.append(d)
        elif char == 'E' and e not in spec_list:
            spec_list.append(e)
    # check for aux spec 'F'
    if dwg_parts[1][2] == 'F' and f not in spec_list:
        spec_list.append(f)

    return spec_list


def controls(dwg_suffix, control_pt, control_size, signal):
    """
    Identify if control type 1 or 2 called for by drawing name. If called out, add correct control part for that control
    type based on signal and control part number indicated on engineered schedule.

    :param str dwg_suffix: accessory/alt/add portion of the drawing code
    :param str control_pt: control part number
    :param str control_size: control size
    :param str signal: signal from signal cell in this particular row from schedule
    :return: controls_list - a list of literal strings indicating spec sheet names for controls parts
    :rtype: list
    """
    controls_list: list[str] = []

    # if control type 1 is found called out in dwg_suffix (controls_list)
    if '+CTRL_1' in dwg_suffix:
        if control_pt.startswith('R'):
            controls_list.extend(['R.pdf', ctrl_1_sm_part[signal]])
        elif control_pt.startswith('Y'):
            controls_list.extend(['Y.pdf', ctrl_1_lg_part[signal][control_size]])
    # if control type 2 is found called out in dwg_suffix (controls_list)
    elif '+CTRL_2' in dwg_suffix:
        if control_pt.startswith('S'):
            controls_list.extend(['S.pdf', ctrl_2_sm_part[signal]])
        elif control_pt.startswith('T'):
            controls_list.extend(['T.pdf', ctrl_2_lg_part[signal]])

    return list(set(controls_list))
