# scripts/smartsheet_update.py
"""
author: Sage Gendron
These functions handle data transfer from the Excel-based estimating process and Smartsheet servers (which is primarily
used as a data repository and lead tracking for sales representative).

update_smartsheet() and mark_as_won() called via the Excel project files by the Estimating and Customer Service
departments respectively.
"""
import datetime
import smartsheet
import xlwings as xw

from smartsheet_utils.create_objects import get_ss_client, create_cell
from smartsheet_utils.upload import upload_attachments, upload_discussions, upload_row_info, sheet_id
from utils.rename import rename

# general Excel cell locations
sch_row_id_cell: str = 'D2'
so_cell: str = 'G22'


def update_ss():
    """
    Updates the primary Estimating Smartsheet with information from the calling Excel file as detailed by created
    cells in this function. Also checks for existing attachments and discussions to update in lieu of create
    (where applicable).

    :return: None
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()
    # instantiate smartsheet client with Quotes smartsheet
    ss_c: smartsheet.Smartsheet = get_ss_client(wb)

    # grab Smartsheet row ID from project file
    job_row_id: int = int(wb.sheets['SCHEDULE'].range(sch_row_id_cell).value)
    # build row/cells with information and push row update to Smartsheet server
    _updated_row = upload_row_info(ss_c, wb, job_row_id)

    # rebuild file names/paths with all possible documents so any that were generated can be uploaded
    quote = rename(wb, 'QUOTE', 'pdf')
    schedule = rename(wb, 'SCHEDULE', 'xlsx')
    submittal = rename(wb, 'SUBMITTAL', 'pdf')
    opt_quote = rename(wb, 'EXCEL QUOTE', 'xlsx')
    # push attachments to Smartsheet server
    _updated_attachments = upload_attachments(ss_c, job_row_id, quote, schedule, submittal, opt_quote)

    # identify filepath for entry as a comment for reference by sales reps
    fpath = '\\'.join(wb.fullname.split('\\')[:-1])
    # push discussions to Smartsheet server
    _updated_discussions = upload_discussions(ss_c, job_row_id, fpath)


def mark_as_won():
    """
    Changes the Quote Status of the active quote's row ID to 'Won', changes the date won to today, and adds the SO#
    to the SO# column.

    :return: None
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()
    schedule = wb.sheets['SCHEDULE']

    # instantiate smartsheet client with Quotes smartsheet with api key selected from initials in filename
    ss_c: smartsheet.Smartsheet = get_ss_client(wb)

    # grab SO# from cell in schedule
    so_no: str = schedule.range(so_cell).value
    if type(so_no) is None or type(so_no) is float:
        raise Exception('Sales Order number not entered. Please add and try again.')

    # build SS row object using the job row ID entered into Excel file
    job_row_id: int = int(schedule.range(sch_row_id_cell).value)
    job_row: smartsheet.Smartsheet.models.Row = ss_c.models.Row()
    job_row.id = job_row_id

    # create SS cell instance, assign date won (always today), add cell to Row
    job_row.cells.append(create_cell(ss_c, 'Date Won', cell_val=datetime.date.today().strftime('%Y-%m-%d')))
    # create SS cell instance, assign sales order # (NoneType if cell is blank), add cell to Row
    job_row.cells.append(create_cell(ss_c, 'SO #', cell_val=so_no))
    # create SS cell instance, assign quote status information, add cell to Row
    job_row.cells.append(create_cell(ss_c, 'Status', cell_val='Won'))

    # push SS row object as an update to SS server via previously instantiated Sheets object
    _updated_row = ss_c.Sheets.update_rows(sheet_id, [job_row])
