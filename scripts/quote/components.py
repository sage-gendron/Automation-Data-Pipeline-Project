# components.py
"""
author: Sage Gendron
Uses information provided to return a base model/part number to be altered by required connection types.
"""


def typ_component(rate, size, char):
    """
    Creates base pn depending on component letter. For 2" and below, non-special_case_1 only.

    :param float rate: flow rate for the given kit
    :param str size: system size (in letter form)
    :param str char: indexed character in drawing file representing component to be quoted
    :return:
        - pn - part number to be altered
    :rtype: str
    """
    pn: str = ''

    if char == 'A':
        # check for small size package
        if size == 'B' and 0 < rate <= 5:
            pn = f"ASM-{size}++"
        else:
            pn = f"A-{size}++"
    elif char == 'B':
        pn = f"B-{size}++"
    elif char == 'C':
        pn = f"C-{size}++"
    elif char == 'D':
        # check for small size package
        if size == 'B' and 0 < rate <= 5:
            pn = f"DSM-{size}++"
        else:
            pn = f"D-{size}++"
    elif char == 'E':
        pn = f"E-{size}FF"

    return pn


def sp_case_1_component(rate, size, char):
    """
    Creates base pn with ss trim depending on component letter. For 2" and below only.

    :param float rate: rate for the given kit
    :param str size: system size (in letter form)
    :param str char: indexed character in drawing file representing component to be quoted
    :return:
        - pn - base part number to be altered based on connection types
    :rtype: str
    """
    pn: str = ''

    if char == 'A':
        # check for small size package
        if size == 'B' and 0 < rate <= 5:
            pn = f"ASMSP1-{size}++"
        else:
            pn = f"ASP1-{size}++"
    elif char == 'B':
        pn = f"BSP1-{size}++"
    elif char == 'C':
        pn = f"CSP1-{size}++"
    elif char == 'D':
        # check for small size package
        if size == 'B' and 0 < rate <= 5:
            pn = f"DSMSP1-{size}++"
        else:
            pn = f"DSP1-{size}++"
    elif char == 'E':
        pn = f"ESP1-{size}FF"

    return pn


def lg_component(rate, size, char):
    """
    Creates base part number depending on component letter. For large size packages only.

    :param float rate: rate for the given package
    :param str size: system size (in letter form)
    :param str char: indexed character in drawing file representing component to be quoted
    :return: pn - base part number to be altered based on connection types
    :rtype: str
    """
    pn: str = ''

    if char == 'A':
        pn = f"A-{size}LG"
    elif char == 'B':
        pn = f"B-{size}LG"
    elif char == 'C':
        pn = f"C-{size}LG"
        if size == 'G':
            if float(rate) < 50:
                pn = f"{pn}-LOW"
            else:
                pn = f"{pn}-HIGH"
        elif size == 'H':
            if float(rate) < 100:
                pn = f"{pn}-LOW"
            else:
                pn = f"{pn}-HIGH"
    elif char == 'D':
        pn = f"D-{size}LG"
        if size == 'G':
            if float(rate) < 65:
                pn = f"{pn}-LOW"
            else:
                pn = f"{pn}-HIGH"
        elif size == 'H':
            if float(rate) < 120:
                pn = f"{pn}-LOW"
            else:
                pn = f"{pn}-HIGH"
    elif char == 'E':
        pn = f"E-{size}LG"

    return pn
