# list_dwgs.py
import pandas as pd
import xlwings as xw
import os
"""
author: Sage Gendron
Read all .pdf files in a directory to an excel file by hierarchically column.
Copies data into an excel file and the quote generation template file to be referenced in drop down menus for automated 
estimating.
"""

# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
copy_area = 'B1:ZZ50'
folder_loc = r'C:\Estimating\CAD Drawings'
list_dwgs_loc = r'C:\Estimating\Data\list_dwgs.xlsx'
tag_temp_loc = r'C:\Estimating\Customer\JOB_WIP 20220406.xlsm'


def list_files():
    """
    Crawl through drawing directories to find drawings and folder names to be listed in project template.

    :returns: dwg_list (:py:class:'list') - data structure containing dwg files, pdf files, and the paths to get to them
    :rtype: list
    """
    # instantiate list variable to be placed into list_dwgs file with column names
    dwg_list: list[list[str, str, str]] = [['path', 'dwg', 'pdf']]

    # walk through subdirectories to find drawings to be listed for submittal pull and quote production
    path: str
    subdir: list[str]
    files: list[str]
    for path, subdir, files in os.walk(folder_loc):
        # skip folders as follows
        if '_archive' in path.lower():
            continue
        if '_edgecase' in path.lower():
            continue
        if 'engineer-specific' in path.lower():
            continue
        # if folder not skipped...
        for file in files:
            # only grab files with .pdf filetypes
            if file.endswith('.pdf'):
                # add the file to the list to be placed into the list_dwgs file
                dwg_list.append([path, f"{file[:-4]}.dwg", file])

    return dwg_list


def kit_type_by_column(lst):
    """
    Transforms the drawings in list of lists format to dictionary format, sends it to a DataFrame, sends to an Excel
    file with columns as folder names, copies to project template file, saves, and closes all files.

    :param list lst:
    :return: None
    """
    columns: list[str] = []
    some_dict: dict[str, list[str]] = {}

    # transform the list of lists into a dictionary, so it can be ordered and transposed into Excel
    row: list[str]
    for row in lst:
        path: list[str] = row[0].split('\\')
        if path[-1] not in some_dict.keys():
            columns.append(path[-1])
            some_dict[path[-1]] = [row[2]]
        else:
            some_dict[path[-1]].append(row[2])

    # create a dataframe and transpose so that the folders are headers and drawings below respective folders
    df_raw: pd.DataFrame = pd.DataFrame.from_dict(some_dict, orient='index').transpose()

    # send the dataframe to excel as the list_dwgs.xlsx file
    df_raw.to_excel(list_dwgs_loc, sheet_name='list_dwgs', index=False, columns=columns)

    # copy list_dwgs excel file to schedule data sheets
    wb_list_dwgs: xw.Book = xw.Book(list_dwgs_loc)
    wb_template_combo: xw.Book = xw.Book(tag_temp_loc)

    # copy list_dwgs sheet to COMBO doc, save, and exit the file
    wb_list_dwgs.sheets['list_dwgs'].range(copy_area).copy(wb_template_combo.sheets['list_dwgs'].range(copy_area))
    wb_template_combo.save()
    wb_template_combo.app.quit()


if __name__ == '__main__':
    # crawl folders to generate a list of filepaths and drawing names that met criteria
    file_list = list_files()
    # transpose and place the lists into an Excel file to be copied into the project template
    kit_type_by_column(file_list)
