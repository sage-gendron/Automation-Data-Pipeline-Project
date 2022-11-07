import xlwings as xw
import os
import shutil
from rename import rename

# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
original_quote: str = r'C:\Estimating\Customer\Quote Template.xlsx'
original_flat_schedule: str = r'C:\Estimating\Customer\Schedule Template.xlsx'
to_orders: str = r'C:\Production\2022\New Orders'

internal_quote_range: str = 'E2:L682'
flat_quote_range: str = 'C2:J682'

job_info_range: str = 'A13:S26'
general_system_range: str = 'B33:L1032'
cv_info_range: str = 'Q33:T1032'
cv2_info_range: str = 'W33:X1032'
smart_notes_col: str = 'AC28:AC1032'
flat_notes_col: str = 'Y28:Y1032'

balance_formula_range: str = 'M33:P1032'
cv_formula_range: str = 'U33:V1032'

sales_order: str = 'G22'
job_name_cell: str = 'C19'


def copy_customer_quote_file(wb):
    """
    author: Sage Gendron \n
    Copies blank flat quote file to calling workbook directory with correct naming scheme.

    :param xlwings.Book wb: calling filepath/filename to be modified
    :return: target (:py:class:'str') - filepath/filename of copied blank flat quote
    :rtype: str
    """
    # generate new filepath + filename
    target: str = rename(wb, 'QUOTE SALES', 'xlsx')

    # copy file from original location to location function being called from
    shutil.copyfile(original_quote, target)

    return target


def generate_customer_quote():
    """
    author: Sage Gendron \n
    Calls copy flat quote file and copies active quote information to flat quote file.

    :return: None
    """
    wb: xw.Book = xw.Book.caller()

    # Copy flat quote file and create xw instance for quote
    flat_quote_fname: str = copy_flat_quote_file(wb)
    wb_fq: xw.Book = xw.Book(flat_quote_fname)

    # copy only active quote cell text (not formulas) to new flat quote file
    wb.sheets['QUOTE'].range(internal_quote_range).copy(wb_fq.sheets['Quote'].range(flat_quote_range))

    # save file, but DO NOT close, so flat quote print area can be adjusted
    wb_fq.save()


def copy_customer_schedule_file(wb):
    """
    author: Sage Gendron \n
    Copies blank flat schedule file to same directory as the calling workbook with correct naming scheme.

    :param xlwings.Book wb: as a parameter, calling workbook filepath
    :return: target (:py:class:'str') - filepath of newly copied flat schedule file
    :rtype: str
    """
    # manipulate calling workbook filepath to identify new flat schedule filepath/name
    target: str = rename(wb, 'SCHEDULE', 'xlsx')

    # copy flat schedule template to target location
    try:
        shutil.copyfile(original_flat_schedule, target)
    except PermissionError:
        raise Exception('Please close the existing flat schedule file and try again.')

    return target


def generate_customer_schedule():
    """
    author: Sage Gendron \n
    Copies scheduled information and formulas to a flat schedule file. Saves the file in the current folder,
    named accordingly.

    :return: None - saves file in folder
    """
    wb: xw.Book = xw.Book.caller()
    schedule = wb.sheets['SCHEDULE']

    # Copy flat quote file and create xw instance for quote
    customer_schedule: str = copy_flat_schedule_file(wb)
    wb_fs: xw.Book = xw.Book(customer_schedule)

    # Copy text only fields
    schedule.range(job_info_range).copy(wb_fs.sheets['SCHEDULE'].range(job_info_range))  # header info
    schedule.range(general_system_range).copy(wb_fs.sheets['SCHEDULE'].range(general_system_range))  # qty, eq_type, tag, gpm, pkg_key, bv_type, sizes/connection types, bv_model
    schedule.range(cv_info_range).copy(wb_fs.sheets['SCHEDULE'].range(cv_info_range))  # control valve info
    schedule.range(cv2_info_range).copy(wb_fs.sheets['SCHEDULE'].range(cv2_info_range))  # control valve/actuator assembly info
    schedule.range(smart_notes_col).copy(wb_fs.sheets['SCHEDULE'].range(flat_notes_col))  # notes

    # Copy formula fields
    fx1 = schedule.range(balance_formula_range).formula  # balancing formulas
    wb_fs.sheets['SCHEDULE'].range(balance_formula_range).formula = fx1
    fx2 = schedule.range(cv_formula_range).formula  # control valve formulas
    wb_fs.sheets['SCHEDULE'].range(cv_formula_range).formula = fx2

    # save file, but DO NOT close, so schedule can be filtered
    wb_fs.save()


def csr_file_copy():
    """
    author: Sage Gendron \n
    Creates folder in to_orders location and copies the flat schedule, submittal, and tagging schedule for warehouse \n

    :return: None
    """
    wb: xw.Book = xw.Book.caller()
    schedule = wb.sheets['SCHEDULE']

    # grab SO# and job name from schedule
    so_no: str = str(schedule.range(sales_order).value)
    job_name: str = schedule.range(job_name_cell).value

    # check to be sure SO# was populated, else error
    if type(so_no) is None or type(so_no) is float:
        raise Exception('Please enter SO number in the appropriate field prior to copying files.')

    # create folder name/location by concatenating SO# - job name
    so_no_jn: str = f"{so_no} - {job_name}"
    so_no_to_orders: str = f"{to_orders}\\{so_no_jn}"

    # manipulate calling workbook filepath to find submittal
    submittal: str = rename(wb, 'SUBMITTAL', 'pdf')
    sub_fname: str = submittal.split('\\')[-1]

    # manipulate calling workbook filepath to find schedule
    flat_schedule: str = rename(wb, 'SCHEDULE', 'xlsx')

    # create instance of the flat schedule and open
    try:
        wb_fs: xw.Book = xw.Book(flat_schedule)
    except FileNotFoundError:
        raise Exception('Generate flat schedule and try again.')

    # instantiate target file variables
    target_to_orders: str = f"{so_no_to_orders}\\{so_no_jn}.xlsx"
    target_submittal: str = f"{so_no_to_orders}\\{sub_fname}"
    tag_sch_loc: str = wb.fullname.split('\\')
    target_tag_sch: str = '\\'.join(tag_sch_loc[:-1]) + '\\' + so_no_jn + '.xlsx'
    target_boxing: str = f"{so_no_to_orders}\\BOXING SCHEDULE.xlsx"

    # create the folder in to_orders if doesn't already exist
    if not os.path.isdir(so_no_to_orders):
        os.makedirs(so_no_to_orders)

    # copy boxing schedule out of flat schedule into a separate file and save; file stays open on screen
    wb_fs.sheets['PRODUCTION SCHEDULE'].api.Copy()
    xw.books.active.save(target_boxing)

    # copy files from calling folder to newly created job folder in to_orders
    try:
        shutil.copyfile(flat_schedule, target_to_orders)
    except FileNotFoundError:
        raise Exception('Please generate the flat schedule and copy files again.')
    shutil.copyfile(submittal, target_submittal)
    shutil.copyfile(flat_schedule, target_tag_sch)
