# scripts/salesorder/assign.py
"""
author: Sage Gendron
Assigns engineered components extracted from the schedule to each package with summed quantities based on package keys.
"""


def engineered_components(engr_components, part_dict, qty_dict, price_dict):
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
