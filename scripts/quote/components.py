# components.py
"""
author: Sage Gendron

"""


def typ_component(rate, pipe_size, pipe_type, char, is_tgv=False, ta_size=None):
    """
    Creates base pn depending on component letter. For 2" and below, non-special_case_1 only.

    :param float rate: flow rate for the given kit
    :param str pipe_size: runout pipe size (in letter form)
    :param str pipe_type: runout type value from cell in this particular row from the schedule
    :param str char: indexed character in drawing file representing component
    :param bool is_tgv: checks if component is a TGV for handling with isolation valves/cvs
    :param ta_size: list of additional components required if TA different size than kit
    :return:
        - pn - part number to be altered
        - is_tgv - checks if component is a TGV for handling with isolation valves/cvs
        - ta_size - list of additional components required if TA different size than kit
    :rtype: (str, bool, list)
    """
    pn: str = None

    # if TY
    if char == 'Y':
        # check for compact criteria
        if pipe_size == 'B' and 0 < rate <= 3.1:
            pn = f"TY1-{pipe_size}++"
        else:
            pn = f"TY-{pipe_size}++"
    # if TU
    elif char == 'U':
        pn = f"TU-{pipe_size}++"
    # if TB
    elif char == 'B':
        pn = f"TB-{pipe_size}++"
    # if TA
    elif char == 'A':
        # check for compact criteria
        if pipe_size == 'B' and 0 < rate <= 3.1:
            pn = f"TA1-{pipe_size}++-L"
        # check for 1-1/4" kit, but 1" TA
        elif pipe_size == 'D' and 0 < rate < 8.5:
            pn = 'TA-CF+-L'
            ta_size = [['CN-C', 'H119-QM', 'TL-C'], [1, 1, 2]]
        # check for 1-1/2" kit, but 1" TA
        elif pipe_size == 'E' and 0 < rate < 8.5:
            pn = 'TA-CF+-L'
            ta_size = [['CN-C', 'H119-RM', 'TL-C'], [1, 1, 2]]
        # check for 1" kit, 1-1/4" TA
        elif pipe_size == 'C' and rate > 12:
            pn = 'TA-DF+-L'
            ta_size = [['HB-DC', 'TL-D'], [1, 1]]
        # check for 2" kit, 1-1/2" TA
        elif pipe_size == 'F' and 0 < rate < 19:
            pn = 'TA-EF+-L'
            ta_size = [['CN-E', 'H119-SR', 'TL-E'], [1, 1, 2]]
        # check for 1-1/2" kit, 2" TA
        elif pipe_size == 'E' and rate > 26:
            pn = 'TA-FF+-L'
            ta_size = [['HB-FE', 'TL-F'], [1, 1]]
        else:
            pn = f"TA-{pipe_size}++-L"
        # check if pipe type is sweat and TA size was changed so SWT adapter and labor can be added
        if pipe_type == 'SWT' and ta_size:
            ta_size[0].extend([f"CA-{pipe_size}", f"TL-{pipe_size}"])
            ta_size[1].extend([1, 1])
    # if TGV, but also set is_tgv to true for future reference (+IR, +IRL, etc.)
    elif char == 'G':
        pn = f"TGV-{pipe_size}FF"
        is_tgv = True
    # if NT
    elif char == 'N':
        pn = f"NT-{pipe_size}++"

    return pn, is_tgv, ta_size


def sp_case_1_component(rate, size, sys_type, char):
    """
    Creates base pn with ss trim depending on component letter. For 2" and below only.

    :param float rate: flow rate for the given kit
    :param str size: runout pipe size (in letter form)
    :param str sys_type: runout type value from cell in this particular row from the schedule
    :param str char: indexed character in drawing file representing component
    :return:
        - pn - base part number to be altered based on connection types
        - is_tgv - checks if component is a TGV for handling with isolation valves/cvs
        - ta_size - list of additional components required if TA different size than kit
    :rtype: (str, bool, list)
    """
    pn: str = None

    # if TY
    if char == 'Y':
        # check for compact criteria
        if size == 'B' and 0 < rate <= 3.1:
            pn = f"TY1SS-{size}++"
        else:
            pn = f"TYSS-{size}++"
    # if TU
    elif char == 'U':
        pn = f"TU-{size}++"
    # if TB
    elif char == 'B':
        pn = f"TBSS-{size}++"
    # if TA
    elif char == 'A':
        # check for compact criteria
        if size == 'B' and 0 < rate <= 3.1:
            pn = f"TA1SS-{size}++-L"
        # check for 1-1/4" kit, but 1" TA
        elif size == 'D' and 0.0 < rate < 8.5:
            pn = 'TASS-CF+-L'
            ta_size = [['CN-C', 'H119-QM', 'TL-C'], [1, 1, 2]]
        # check for 1-1/2" kit, but 1" TA
        elif size == 'E' and 0.0 < rate < 8.5:
            pn = 'TASS-CF+-L'
            ta_size = [['CN-C', 'H119-RM', 'TL-C'], [1, 1, 2]]
        # check for 1" kit, 1-1/4" TA
        elif size == 'C' and rate > 12:
            pn = 'TA-DF+-L'
            ta_size = [['CN-D', 'HB-ED', '100-107SSG', 'TL-D', 'TL-E'], [1, 1, 1, 2, 1]]
        # check for 2" kit, 1-1/2" TA
        elif size == 'F' and 0 < rate < 20:
            pn = 'TA-EF+-L'
            ta_size = [['CN-E', 'HB-FE', '100-108SSG', 'TL-E', 'TL-F'], [1, 1, 1, 2, 1]]
        # check for 1-1/2" kit, 2" TA
        elif size == 'E' and rate > 26:
            pn = 'TA-FF+-L'
            ta_size = [['HB-FE', 'CN-E', '100-107SSG', 'TL-E', 'TL-F'], [1, 1, 1, 2, 1]]
        else:
            pn = f"TASS-{size}++-L"

        if sys_type == 'SWT' and ta_size:
            ta_size[0].extend([f"CA-{size}", f"TL-{size}"])
            ta_size[1].extend([1, 1])
    # if TGV, but also set is_tgv to true for future reference (+IR, +IRL, etc.)
    elif char == 'G':
        pn = f"TGV-{size}FF"
        is_tgv = True
    # if NT
    elif char == 'N':
        pn = f"NTSS-{size}++"

    return pn


def lg_component(rate, size, char):
    """
    Creates base pn depending on component letter. For 2-1/2" and above only.

    :param float rate: flow rate for the given kit
    :param str size: runout pipe size (in letter form)
    :param str char: indexed character in drawing file representing component
    :return: pn - base part number to be altered based on connection types
    :rtype: str
    """
    pn: str = ''

    # if TS
    if char == 'Y':
        pn = f"TS-{size}LF"
    # if TB
    elif char == 'B':
        pn = f"TB-{size}LF"
        if size == 'G':
            if float(rate) < 64.0:
                pn = f"{pn}-L"
            else:
                pn = f"{pn}-H"
        elif size == 'H':
            if float(rate) < 101.0:
                pn = f"{pn}-L"
            else:
                pn = f"{pn}-H"
    # if TA
    elif char == 'A':
        pn = f"TA-{size}W"
        if size == 'G':
            if float(rate) <= 80.0:
                pn = f"{pn}-L"
            else:
                pn = f"{pn}-H"
        elif size == 'H':
            if float(rate) <= 135.0:
                pn = f"{pn}-L"
            else:
                pn = f"{pn}-H"
        elif size == 'I':
            if float(rate) <= 270.0:
                pn = f"{pn}-L"
            else:
                pn = f"{pn}-H"
        elif size == 'J':
            if float(rate) <= 540.0:
                pn = f"{pn}-L"
            else:
                pn = f"{pn}-H"
        elif size == 'K':
            if float(rate) <= 945.0:
                pn = f"{pn}-L"
            else:
                pn = f"{pn}-H"
    # if TGV
    elif char == 'G':
        pn = f"TGV-{size}FF"
    # if BFV
    elif char == 'I':
        pn = f"BF-{size}LL"

    return pn
