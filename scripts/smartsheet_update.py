# smartsheet_update.py
import datetime
import json
import os
import smartsheet
import xlwings as xw
"""
author: Sage Gendron

"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
# Smartsheet API keys by user
sg_api_key: str = '###'
user2_api_key: str = '###'
user3_api_key: str = '###'
user4_api_key: str = '###'

# Unique ID for estimating Smartsheet sheet
sheet_id: str = '@@@'

# create important, unique column ID dictionary for reference below
column_ids: dict[str, int] = {
    'Quote Status': 111000, 'Quoted By': 222000, 'Institution': 333000,
    'Buy Sell Rep Name': 444000, 'Opportunity': 555000, 'Multiplier': 666000,
    'Deal Total': 777000, 'Quote Request Contact': 888000, 'Regional': 999000,
    'Bid Date': 101010, 'Date of Quote': 111111, 'Closed Date': 121212,
    'Deal Status': 131313, 'Loss Reason': 141414, 'Notes': 151515,
    'Industry': 161616, 'Balance Valve Type': 171717, 'SS': 181818,
    'Equipment Type': 191919, 'Mechanical Engineer': 202020, 'Address': 212121,
    'City': 222222, 'State': 232323, 'Zip': 242424, 'Contractor': 252525,
    'SO #': 262626, 'Phase': 272727
}

# Project Excel file cell variables
# for iterated updating
simple_cols: dict[str, str] = {
    'Industry': 'D3', 'Balance Valve Type': 'D4', 'Equipment Type': 'D6', 'Mechanical Engineer': 'C22',
    'Contractor': 'C21', 'Address': 'D7', 'City': 'D8', 'State': 'D9', 'Zip': 'D10', 'Phase': 'G24'
}
# for logical updating
_sch_job_name_cell: str = 'C19'  # not currently implemented
_sch_rep_cell: str = 'C23'  # not currently implemented
sch_owner_cell: str = 'G23'
sch_contact_cell: str = 'C24'
sch_row_id_cell: str = 'D2'
sch_ss_cell: str = 'D5'
q_x_cell: str = 'A3'
q_total_cell: str = 'AA4'
so_cell: str = 'G22'


def get_ss_client(wb):
    """
    Uses filepath from xw.Book to identify author and create a Smartsheet client using their API key.

    :param xw.Book wb: Excel workbook to grab the estimator's initials from
    :returns ss_c: A smartsheet client object created with the estimator's API key
    :rtype: smartsheet.Smartsheet
    """
    # find the file name by splitting the filepath of the workbook
    jn = wb.fullname.split('\\')[-1].split('_')
    # look for initials for quoting personnel to select the correct API key with which to instantiate the client
    ss_c: smartsheet.Smartsheet
    if 'USER3' in jn:
        ss_c = smartsheet.Smartsheet(user3_api_key)
    elif 'USER2' in jn:
        ss_c = smartsheet.Smartsheet(user2_api_key)
    elif 'USER4' in jn:
        ss_c = smartsheet.Smartsheet(user4_api_key)
    else:
        ss_c = smartsheet.Smartsheet(sg_api_key)
    ss_c.errors_as_exceptions(True)

    return ss_c


def create_cell(ss_c, col_name, wb=None, sheet=None, xl_cell=None, is_float=False, is_bool=False, cell_val=None):
    """
    Creates a new Smartsheet Cell object, passes values as required (or pulls them from a workbook), and returns
    the cell to be appended to a Row object.

    :param smartsheet.Smartsheet ss_c: Smartsheet client object
    :param str col_name: name of the column for reference from the column_ids global variable
    :param xw.Book wb: Excel workbook from which to extract data
    :param str sheet: Sheet name to reference when extracting from wb (xlwings workbook)
    :param str xl_cell: Cell reference from which to extract data within the sheet, book above
    :param bool is_float: indicates if value should be read as a float
    :param bool is_bool: indicates if values should be read as a bool
    :param cell_val: cell value, if known, so it can be attributed to the new cell object
    :returns new_cell: a new cell object with value as specified
    :rtype: smartsheet.Smartsheet.models.Cell
    """
    # create the new Cell object
    new_cell: smartsheet.Smartsheet.models.Cell = ss_c.models.Cell()
    # assign the column id to the cell object, so it has a place to be updated in the row
    new_cell.column_id = column_ids[col_name]

    # assign the cells value based on if the value was passed to this function or if it needs to be pulled from excel
    if cell_val:
        new_cell.value = cell_val
    elif is_float:
        new_cell.value = float(wb.sheets[sheet.upper()].range(xl_cell).value)
    elif is_bool:
        new_cell.value = bool(wb.sheets[sheet.upper()].range(xl_cell).value)
    else:
        new_cell.value = wb.sheets[sheet.upper()].range(xl_cell).value

    new_cell.strict = False

    return new_cell


def update_ss():
    """
    Updates the Quotes Smartsheet with information from the calling project file as detailed by created cells in this
    function. Also checks for existing attachments and discussions to update in lieu of create (where applicable).

    :return: None
    """
    wb: xw.Book = xw.Book.caller()

    quote_name = wb.fullname.split('\\')
    quote_name[-1] = f"{quote_name[-1][:-4]}xlsm"
    quote_fname = quote_name[-1].split('_')
    quote_sales_fname = quote_name[-1].split('_')
    quote_fname[-2] = 'QUOTE'
    quote_sales_fname[-2] = 'QUOTE SALES'
    quote_fname = '_'.join(quote_fname)
    quote_sales_fname = '_'.join(quote_sales_fname)
    quote_sales_fname = f"{quote_sales_fname[:-1]}x"
    quote_name = quote_name[:-1]
    quote_sales = quote_name
    fpath = '\\'.join(quote_name)
    quote_name = '\\'.join(quote_name) + f"\\{quote_fname}"
    quote_sales = '\\'.join(quote_sales) + f"\\{quote_sales_fname}"

    submittal_name = wb.fullname.split('\\')
    submittal_name[-1] = f"{submittal_name[-1][:-4]}pdf"
    submittal_fname = submittal_name[-1].split('_')
    submittal_fname[-2] = 'SUBMITTAL'
    submittal_fname = '_'.join(submittal_fname)
    submittal_name = submittal_name[:-1]
    submittal_name = '\\'.join(submittal_name) + '\\' + submittal_fname

    schedule_name = wb.fullname.split('\\')
    schedule_name[-1] = f"{schedule_name[-1][:-1]}x"
    schedule_fname = schedule_name[-1].split('_')
    schedule_fname[-2] = 'SCHEDULE'
    schedule_fname = '_'.join(schedule_fname)
    schedule_name = schedule_name[:-1]
    schedule_name = '\\'.join(schedule_name) + '\\' + schedule_fname

    # instantiate smartsheet client with Quotes smartsheet
    ss_c: smartsheet.Smartsheet = get_ss_client(wb)

    # build SS row object and append cells for info push
    job_row_id: int = int(wb.sheets['SCHEDULE'].range(sch_row_id_cell).value)
    job_row: smartsheet.Smartsheet.models.Row = ss_c.models.Row()
    job_row.id = job_row_id

    # build comment object with filepath for text
    fpath_comment: smartsheet.Smartsheet.models.Comment = ss_c.models.Comment({'text': fpath})

    # instantiate flags to skip update calls below
    skip_contact: bool = False
    skip_institution: bool = False
    update_comment = None
    update_quote_upload = None
    update_schedule_upload = None
    update_submittal_upload = None
    update_quote_sales_upload = None

    # grab existing row, attachments, discussions to help decide what to update
    current_row: smartsheet.Smartsheet.models.Row = ss_c.Sheets.get_row(sheet_id, job_row_id, include='columns')
    current_attachments = ss_c.Attachments.list_row_attachments(sheet_id, job_row_id)
    current_discussions = ss_c.Discussions.get_row_discussions(sheet_id, job_row_id, include_all=True)

    # iterate through existing information in row and set IDs to update if required
    cell: smartsheet.Smartsheet.models.Cell
    for cell in current_row.cells:
        if cell.column_id == column_ids['Quote Request Contact'] and cell.value not in ('', None):
            skip_contact = True
        if cell.column_id == column_ids['Institution'] and cell.value not in ('', None):
            skip_institution = True

    discussion: smartsheet.Smartsheet.models.Discussion
    for discussion in current_discussions.data:
        for comment in discussion.comments:
            if comment.text == fpath:  # this isn't getting triggered? might be str type issue
                update_comment = comment.id

    attachment: smartsheet.Smartsheet.models.Attachment
    for attachment in current_attachments.data:
        if attachment.name == quote_fname[:-4] + 'pdf':
            update_quote_upload = attachment.id
        elif attachment.name == schedule_fname[:-1] + 'x':
            update_schedule_upload = attachment.id
        elif attachment.name == submittal_fname:
            update_submittal_upload = attachment.id
        elif attachment.name == quote_sales_fname:
            update_quote_sales_upload = attachment.id

    # create SS cell instance, assign quote status to Quote Done, add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Quote Status', cell_val='Quote Done'))

    # create SS cell instance, assign multiplier from quote, add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Multiplier', xl_cell=q_x_cell, wb=wb, sheet='QUOTE', is_float=True))

    # create SS cell instance, assign deal total from quote, add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Deal Total', xl_cell=q_total_cell, wb=wb, sheet='QUOTE', is_float=True))

    # create SS cell instance, assign SS boolean value, add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'SS', wb=wb, sheet='SCHEDULE', xl_cell=sch_ss_cell, is_bool=True))

    if not skip_contact:
        job_row.cells.append(create_cell(ss_c, 'Quote Request Contact', wb=wb, sheet='SCHEDULE',
                                         xl_cell=sch_contact_cell))

    if not skip_institution and type(cell_val := wb.sheets['SCHEDULE'].range(sch_owner_cell).value) is str:
        job_row.cells.append(create_cell(ss_c, 'Institution', cell_val=cell_val))

    # create SS cell instance, assign quote date (always today), add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Date of Quote', cell_val=datetime.date.today().strftime('%Y-%m-%d')))

    # loop through simple cell updates that will always happen if information is provided
    col: str
    xl_cell: str
    for col, xl_cell in simple_cols.items():
        if type(cell_val := wb.sheets['SCHEDULE'].range(xl_cell).value) is str:
            job_row.cells.append(create_cell(ss_c, col, cell_val=cell_val))

    # push row update to SS server
    _updated_row = ss_c.Sheets.update_rows(sheet_id, [job_row])

    # add filepath as comment
    if update_comment is not None:
        _updated_comment = ss_c.Discussions.update_comment(sheet_id, update_comment, fpath_comment)
    else:
        _updated_comment = ss_c.Discussions.create_discussion_on_row(sheet_id, job_row_id,
                                                                     ss_c.models.Discussion({'comment': fpath_comment})
                                                                     )

    # initiate file attachment upload to row by filename/path
    if update_quote_upload is not None:
        _upload_quote = ss_c.Attachments.attach_new_version(sheet_id, update_quote_upload,
                                                            (f"{quote_fname[:-4]}pdf", open(f"{quote_name[:-4]}pdf",
                                                                                            'rb'), 'application/pdf')
                                                            )
    else:
        _upload_quote = ss_c.Attachments.attach_file_to_row(sheet_id, job_row_id, (f"{quote_fname[:-4]}pdf",
                                                            open(f"{quote_name[:-4]}pdf", 'rb'), 'application/pdf')
                                                            )
    try:
        if update_submittal_upload is not None:
            _upload_submittal = ss_c.Attachments.attach_new_version(sheet_id, update_submittal_upload, (submittal_fname,
                                                                    open(submittal_name, 'rb'), 'application/pdf')
                                                                    )
        else:
            _upload_submittal = ss_c.Attachments.attach_file_to_row(sheet_id, job_row_id, (submittal_fname,
                                                                    open(submittal_name, 'rb'), 'application/pdf')
                                                                    )
    except FileNotFoundError:
        pass
    if update_schedule_upload is not None:
        _upload_schedule = ss_c.Attachments.attach_new_version(sheet_id, update_schedule_upload,
                                                               (f"{schedule_fname[:-1]}x", open(schedule_name, 'rb'),
                                                                'application/vnd.openxmlformats-'
                                                                'officedocument.spreadsheetml.sheet')
                                                               )
    else:
        _upload_schedule = ss_c.Attachments.attach_file_to_row(sheet_id, job_row_id,
                                                               (f"{schedule_fname[:-1]}x", open(schedule_name, 'rb'),
                                                                'application/vnd.openxmlformats-'
                                                                'officedocument.spreadsheetml.sheet')
                                                               )
    if update_quote_sales_upload is not None:
        try:
            _upload_sales_quote = ss_c.Attachments.attach_new_version(sheet_id, update_quote_sales_upload,
                                                                      (quote_sales_fname, open(quote_sales, 'rb'),
                                                                       'application/vnd.openxmlformats-'
                                                                       'officedocument.spreadsheetml.sheet')
                                                                      )
        except FileNotFoundError:
            pass

    else:
        try:
            _upload_sales_quote = ss_c.Attachments.attach_file_to_row(sheet_id, job_row_id, (quote_sales_fname,
                                                                      open(quote_sales, 'rb'),
                                                                      'application/vnd.openxmlformats-'
                                                                      'officedocument.spreadsheetml.sheet')
                                                                      )
        except FileNotFoundError:
            pass


def mark_as_won():
    """
    Changes the deal status of the project file's row ID to 'Won', changes the closed date to today, and adds the SO#
    to the SO# column (all in the Quotes Smartsheet).

    :return: None
    """
    wb: xw.Book = xw.Book.caller()
    schedule = wb.sheets['SCHEDULE']

    # instantiate smartsheet client with Quotes smartsheet with api key selected from initials in filename
    ss_c: smartsheet.Smartsheet = get_ss_client(wb)

    # grab SO# from cell in schedule
    so_no: str = schedule.range(so_cell).value
    if type(so_no) in (None, float):
        raise Exception('SO number not entered. Please add and try again.')

    # build SS row object using the job row ID entered into Excel file
    job_row_id: int = int(schedule.range(sch_row_id_cell).value)
    job_row: smartsheet.Smartsheet.models.Row = ss_c.models.Row()
    job_row.id = job_row_id

    # create SS cell instance, assign deal status information, add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Deal Status', cell_val='Won'))

    # create SS cell instance, assign closed date (always today), add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Closed Date', cell_val=datetime.date.today().strftime('%Y-%m-%d')))

    # create SS cell instance, assign sales order # (NoneType if cell is blank), add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'SO #', cell_val=so_no))

    # push SS row object as an update to SS server via previously instantiated Sheets object
    _updated_row = ss_c.Sheets.update_rows(sheet_id, [job_row])
