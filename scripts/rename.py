# rename.py


def rename(wb, fname, ftype):
    """
    author: Sage Gendron \n
    Rename the filepath for a given Excel Workbook based on the double underscore schema, the fname parameter,
    and the ftype parameter.

    :param xlwings.Book wb: excel file with filename/path to be altered
    :param str fname: new filename to be placed between double underscores
    :param str ftype: new filetype to be placed at the end of the filename (must match file to be exported)
    :return: target (:py:class:'str') - full filepath with new filename/type at the end
    :rtype: str
    """
    # retrieve filepath for the given workbook
    target = wb.fullname
    # split the filepath by \
    target = target.split('\\')
    # split the actual filename by _ assuming the file has used the double underscore schema
    name = target[-1].split('_')
    # check to make sure the filename was split properly so the filename can be correctly renamed
    if len(name) == 1:
        raise Exception('Please ensure your filename is following the "PROJECT_FILENAME_QUOTENUM" schema and try again.')
    # set the filename according to the parameter given
    name[-2] = fname
    # set the filetype from .xlsm to the parameter given
    name[-1] = f"{name[-1][:-4]}{ftype}"
    # reassemble the filename with underscores
    name = '_'.join(name)
    # reassemble the filepath with \ and append the filename at the end
    target = '\\'.join(target[:-1]) + '\\' + name

    return target
