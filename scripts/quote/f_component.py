# scripts/quote/f_component.py
"""
author: Sage Gendron
Quotes F components where required on small size kits, else returns False for f_check.
"""
from connections import size_dict


def sm_f_component(part_list, part_quantity, components, size, sys_type, conn_size, conn_type, f_size, f_type):
    """
    Quotes f components based on information available specifically catches size changes and tries to account for
    connection sizes where possible.

    :param list part_list: list of parts quoted thus far
    :param list part_quantity: list of quantities to pair with part_list entries by index
    :param list components: smart kit code (drawing name) split by hyphens
    :param str size: system size value (in letter form)
    :param str sys_type: system type value
    :param str conn_size: connection size value (in letter form)
    :param str conn_type: connection type value
    :param f_size: f component size (only if needed to be specified in special cases (ie not matching system size))
    :param f_type: f component type (only if needed to be specified so far)
    :return:
        - part_list - list of parts quoted thus far
        - part_quantity - list of quantities to pair with part_list entries by index
        - f_check - were there f components?
    :rtype: (list, list, bool)
    """
    f_check: bool = False
    pn_f_add = None

    if components[1][2] == 'F':
        # if f components on system side of packages, f component size/type to match system size/type
        if '=RF' in components[1]:
            f_size = size
            f_type = sys_type

        # handle if f component size is not provided
        if type(f_size) is float or f_size is None:
            # if connection size not provided, f component size set to system size
            if type(conn_size) is float:
                f_size = size
            # if connection size provided, set f component size equal to connection size
            else:
                try:
                    f_size = size_dict[conn_size]
                except KeyError:
                    raise Exception(f"{size} not a usable size for F components.")

        # if no previously determined f component type and no connection type provided
        if f_type is None and type(conn_type) is float:
            f_type = 'M'
        # if no previously determined f component type and conn_type determined
        elif f_type is None and conn_type not in ('THD', 'TBD'):
            f_type = 'N' if conn_type == 'NON-THD' else 'M'
        # if no previously determined f component type or f component type THD/TBD
        elif type(f_type) is float or f_type in ('THD', 'TBD', 'MALE', 'FEMALE'):
            f_type = 'M'
        # if f component type predetermined to be NON-THD
        elif f_type == 'NON-THD':
            f_type = 'N'
        # if all else fails and f component type is not predetermined, assume male
        else:
            f_type = 'M'

        # add additional parts if special_case_2
        if f_type == 'SP_CASE_2':
            pn_f = f"F-{f_size}MM"
            pn_f_add = [f"SP2-{f_size}F", f"L-{f_size}"]
        else:
            pn_f = f"F-{f_size}M{f_type}"
        # insert f component part number to part list and quantity list
        part_list.insert(1, pn_f)
        part_quantity.insert(1, 2)
        # append pn_f_add to part list and quantity list if exists
        if pn_f_add:
            part_list.extend(pn_f_add)
            for i in range(len(pn_f_add)):
                part_quantity.append(2)
        # add extra aux parts if L in smart kit number alongside F
        if components[1][3] == 'L':
            pn_l = f"L-{f_size}"
            part_list.append(pn_l)
            part_quantity.append(2)
        f_check = True

    return part_list, part_quantity, f_check
