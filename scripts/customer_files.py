# customer_files.py
import xlwings as xw
import os
import shutil
from scripts.utils.rename import rename
"""
author: Sage Gendron
Handles copying template files and copying required information from automated Excel workbooks to 'flat' Excel files for
customer/engineer consumption.
"""

# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
# Template file locations
quote_template: str = r'C:\Estimating\Customer\Quote Template.xlsx'
schedule_template: str = r'C:\Estimating\Customer\Schedule Template.xlsx'

# Production source directory
prod_order_loc: str = r'C:\Production\2022\New Orders'

# Quote pertinent info cell ranges
internal_quote_range: str = 'E2:L682'
customer_quote_range: str = 'C2:J682'

# Schedule pertinent info cell ranges (not Excel formulas)
job_info_range: str = 'A13:S26'
general_system_range: str = 'B33:L1032'
ctrl_info_range: str = 'Q33:T1032'
prod_info_range: str = 'W33:X1032'
smart_notes_col: str = 'AC28:AC1032'
flat_notes_col: str = 'Y28:Y1032'

# Schedule pertinent formula ranges
engr_formula_range: str = 'M33:P1032'
ctrl_formula_range: str = 'U33:V1032'

# Job-specific cells
sales_order_cell: str = 'G22'
job_name_cell: str = 'C19'


def copy_customer_quote_file(wb):
    """
    Copies blank flat quote file to active/calling workbook directory with correct naming scheme.

    :param xlwings.Book wb: calling filepath/filename to be modified
    :return: target - filepath/filename of copied customer quote template
    :rtype: str
    """
    # generate new filepath + filename for the customer quote using the active workbook
    target: str = rename(wb, 'QUOTE SALES', 'xlsx')

    # copy file from original location to location of calling workbook being called from
    shutil.copyfile(quote_template, target)

    return target


def generate_customer_quote():
    """
    Calls function to copy the customer quote template and copies active quote information to newly copied file.

    :return: None
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()

    # copy customer quote template and create Book instance for customer quote
    customer_quote_fname: str = copy_customer_quote_file(wb)
    wb_cq: xw.Book = xw.Book(customer_quote_fname)

    # copy only active quote cell text (not formulas) to new customer quote file
    wb.sheets['QUOTE'].range(internal_quote_range).copy(wb_cq.sheets['QUOTE'].range(customer_quote_range))

    # saves file, but DOES NOT close it, so customer quote print area can be adjusted if required
    wb_cq.save()


def copy_customer_schedule_file(wb):
    """
    Copies customer schedule template to the same directory as the calling workbook.
    This will purposefully throw an error if the file does not use the standardized naming scheme.

    :param xlwings.Book wb: as a parameter, calling workbook filepath
    :return: target - filepath of newly copied flat schedule file
    :rtype: str
    """
    # manipulate calling workbook filepath to identify new flat schedule filepath/name
    target: str = rename(wb, 'SCHEDULE', 'xlsx')

    # copy flat schedule template to target location
    try:
        shutil.copyfile(schedule_template, target)
    except PermissionError:
        raise Exception('Please close the existing flat schedule file and try again.')

    return target


def generate_customer_schedule():
    """
    Copies scheduled information and formulas to a flat schedule file. Saves the file in the current folder,
    named accordingly.

    :return: None - saves file in folder
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()
    smart_schedule = wb.sheets['SCHEDULE']

    # copy flat quote file and create xw instance for quote
    customer_schedule: str = copy_customer_schedule_file(wb)
    wb_fs: xw.Book = xw.Book(customer_schedule)
    dest_schedule = wb_fs.sheets['SCHEDULE']

    # copy text only fields
    # header info
    smart_schedule.range(job_info_range).copy(dest_schedule.range(job_info_range))
    # qty, equipment_type, tag, design_value, pkg_key, product_type, sizes/connection types, product_model
    smart_schedule.range(general_system_range).copy(dest_schedule.range(general_system_range))
    # controls info
    smart_schedule.range(ctrl_info_range).copy(dest_schedule.range(ctrl_info_range))
    # controls production info
    smart_schedule.range(prod_info_range).copy(dest_schedule.range(prod_info_range))
    # notes fields
    smart_schedule.range(smart_notes_col).copy(dest_schedule.range(flat_notes_col))

    # Copy formula only fields
    # engineering formulas
    fx1 = smart_schedule.range(engr_formula_range).formula
    dest_schedule.range(engr_formula_range).formula = fx1
    # controls formulas
    fx2 = smart_schedule.range(ctrl_formula_range).formula
    dest_schedule.range(ctrl_formula_range).formula = fx2

    # save file, but DO NOT close, so schedule can be filtered
    wb_fs.save()


def csr_file_copy():
    """
    Creates folder in prod_orders location and copies the customer schedule, submittal, and packing list for warehouse

    :return: None - Saves files in the production folder denoted by the Sales Order number
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()
    schedule = wb.sheets['SCHEDULE']

    # grab SO# and job name from schedule
    sales_order_num: str = str(schedule.range(sales_order_cell).value)
    job_name: str = schedule.range(job_name_cell).value

    # check to be sure SO# was populated, else error
    if type(sales_order_num) is None or type(sales_order_num) is float:
        raise Exception('Please enter SO number in the appropriate field prior to copying files.')

    # create folder name/location by concatenating SO# - job name
    so_no_jn: str = f"{sales_order_num} - {job_name}"
    so_prod_path: str = f"{prod_order_loc}\\{so_no_jn}"

    # manipulate calling workbook filepath to find submittal
    submittal: str = rename(wb, 'SUBMITTAL', 'pdf')
    submittal_fname: str = submittal.split('\\')[-1]

    # manipulate calling workbook filepath to find schedule
    customer_schedule: str = rename(wb, 'SCHEDULE', 'xlsx')

    # create instance of the flat schedule and open
    try:
        wb_fs: xw.Book = xw.Book(customer_schedule)
    except FileNotFoundError:
        raise Exception('Generate flat schedule and try again.')

    # build target filepaths
    target_to_orders: str = f"{so_prod_path}\\{so_no_jn}.xlsx"
    target_submittal: str = f"{so_prod_path}\\{submittal_fname}"
    tag_sch_loc: str = wb.fullname.split('\\')
    target_prod_sch: str = '\\'.join(tag_sch_loc[:-1]) + f"\\{so_no_jn}.xlsx"
    target_packing: str = f"{so_prod_path}\\PACKING LIST.xlsx"

    # create the folder in to_orders if it doesn't already exist
    if not os.path.isdir(so_prod_path):
        os.makedirs(so_prod_path)

    # copy packing list out of flat schedule into a separate file and save; file stays open on screen
    wb_fs.sheets['PACKING LIST'].api.Copy()
    xw.books.active.save(target_packing)

    # copy files from calling folder to newly created job folder in production queue
    try:
        shutil.copyfile(customer_schedule, target_to_orders)
    except FileNotFoundError:
        raise Exception('Please generate the flat schedule and copy files again.')
    shutil.copyfile(submittal, target_submittal)
    shutil.copyfile(customer_schedule, target_prod_sch)
