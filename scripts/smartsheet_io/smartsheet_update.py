# smartsheet_update.py
import datetime
import smartsheet
import xlwings as xw
from rename import rename
"""
author: Sage Gendron
These functions handle data transfer from the Excel-based estimating process and Smartsheet servers (which is primarily
used as a data repository and lead tracking for sales representative). 

update_smartsheet() and mark_as_won() called via the Excel project files by the Estimating and Customer Service
departments respectively. 
"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
# smartsheet API keys by user
sg_api_key: str = '###'
user2_api_key: str = '###'
user3_api_key: str = '###'
user4_api_key: str = '###'

# unique ID for estimating Smartsheet sheet
sheet_id: str = '@@@'

# create unique column ID dictionary for reference below
column_ids: dict[str, int] = {
    'City': 222222, 'Customer Pricing': 666000, 'Date Won': 121212, 'Engineer': 202020, 'Engineered Component': 171717,
    'Equipment Type': 191919, 'Industry': 161616, 'Phase': 272727, 'Quote Date': 101010, 'Quote Value': 777000,
    'SO #': 262626, 'Special Case 1': 181818, 'State': 232323, 'Status': 111000, 'Street Address': 212121, 'Zip': 242424
}

# general Excel cell locations
sch_row_id_cell: str = 'D2'
so_cell: str = 'G22'  # only used for mark_as_won()

# schedule Excel file cell variables for update_smartsheet()
col_opts: dict[str, str] = {
    'City': 'D8', 'Customer Pricing': 'A3', 'Engineer': 'C22', 'Engineered Component': 'D4', 'Equipment Type': 'D6',
    'Industry': 'D3', 'Phase': 'G24', 'Quote Date': None, 'Quote Value': 'AA4', 'Special Case 1': 'D5', 'State': 'D9',
    'Status': None, 'Street Address': 'D7',  'Zip': 'D10'
}


def get_ss_client(wb):
    """
    Uses filepath from the passed workbook to identify author and create a Smartsheet client using their API key.

    :param xw.Book wb: Excel workbook to grab the estimator's initials from
    :return ss_c: A smartsheet client object created with the estimator's API key
    :rtype: smartsheet.Smartsheet
    """
    # find the file name by splitting the filepath of the workbook
    jn = wb.fullname.split('\\')[-1].split('_')
    # look for initials for quoting personnel to select the correct API key with which to instantiate the client
    ss_c: smartsheet.Smartsheet
    if 'USER2' in jn:
        ss_c = smartsheet.Smartsheet(user3_api_key)
    elif 'USER3' in jn:
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
    :param bool is_bool: indicates if value should be read as a bool
    :param cell_val: cell value, if known/is a stock value, so it can be attributed to the new cell object
    :returns new_cell: a new Cell object with value as specified
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

    # set cell strictness to false; if left to default value (True), doesn't allow server to edit datatypes and can
    # cause silent issues on row update
    new_cell.strict = False

    return new_cell


def upload_attachments(ss_c, row_id, quote, schedule, submittal, opt_quote):
    """
    Checks for already-uploaded attachments (in case Smartsheet is being updated twice for the same row), attaches
    new versions of each file if this is the case. If files were not previously uploaded, uploads new attachments.

    :param smartsheet.Smartsheet ss_c: Smartsheet client object
    :param int row_id: Smartsheet unique row identifying number
    :param str quote: possible file path for generated Quote document (PDF)
    :param str schedule: possible file path for generated Engineered Schedule document (Excel)
    :param str submittal: possible file path for generated submittal document (PDF)
    :param str opt_quote: possible file path for generated Quote document (Excel)
    :return: _updated_attachments - list of Smartsheet server JSON response documents
    :rtype: list
    """
    # grab existing attachments (if they exist) to help decide what to update
    current_attachments = ss_c.Attachments.list_row_attachments(sheet_id, row_id)

    # create flag variables for attachment updates
    update_quote_upload = False
    update_schedule_upload = False
    update_submittal_upload = False
    update_opt_quote_upload = False

    # loop through row attachments to identify if updates are required based on file name (as API directives differ)
    attachment: smartsheet.Smartsheet.models.Attachment
    for attachment in current_attachments.data:
        if attachment.name == quote.split('\\')[-1]:
            update_quote_upload = True
        elif attachment.name == schedule.split('\\')[-1]:
            update_schedule_upload = True
        elif attachment.name == submittal.split('\\')[-1]:
            update_submittal_upload = True
        elif attachment.name == opt_quote.split('\\')[-1]:
            update_opt_quote_upload = True

    # create empty list to store Smartsheet server responses if ever required
    _updated_attachments = []

    # attaches a new version of the pdf quote if one was found, otherwise uploads a new file
    if update_quote_upload:
        _updated_attachments.append(ss_c.Attachments.attach_new_version(sheet_id, update_quote_upload, (
            quote.split('\\')[-1], open(quote, 'rb'), 'application/pdf')))
    else:
        _updated_attachments.append(ss_c.Attachments.attach_file_to_row(sheet_id, row_id, (
            quote.split('\\')[-1], open(quote, 'rb'), 'application/pdf')))

    # attaches a new version of the engineered schedule if one was found, otherwise uploads a new file
    if update_schedule_upload:
        _updated_attachments.append(ss_c.Attachments.attach_new_version(
            sheet_id, update_schedule_upload, (schedule.split('\\')[-1], open(schedule, 'rb'),
                                               'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')))
    else:
        _updated_attachments.append(ss_c.Attachments.attach_file_to_row(
            sheet_id, row_id, (schedule.split('\\')[-1], open(schedule, 'rb'),
                               'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')))

    # attaches a new version of the submittal if one was found, otherwise attaches a new file
    try:
        if update_submittal_upload:
            _updated_attachments.append(ss_c.Attachments.attach_new_version(sheet_id, update_submittal_upload, (
                submittal.split('\\')[-1], open(submittal, 'rb'), 'application/pdf')))
        else:
            _updated_attachments.append(ss_c.Attachments.attach_file_to_row(sheet_id, row_id, (
                submittal.split('\\')[-1], open(submittal, 'rb'), 'application/pdf')))
    # since this file doesn't exist for every quote produced, pass on error
    except FileNotFoundError:
        pass

    # attaches a new version of the Excel customer quote if one was found, otherwise uploads a new file
    try:
        if update_opt_quote_upload:
            _updated_attachments.append(ss_c.Attachments.attach_new_version(
                sheet_id, update_opt_quote_upload, (opt_quote.split('\\')[-1], open(opt_quote, 'rb'),
                                                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')))
        else:
            _updated_attachments.append(ss_c.Attachments.attach_file_to_row(
                sheet_id, row_id, (opt_quote.split('\\')[-1], open(opt_quote, 'rb'),
                                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')))
    # since this file doesn't exist for every quote produced, pass on error
    except FileNotFoundError:
        pass

    return _updated_attachments


def upload_discussions(ss_c, row_id, fpath):
    """
    Checks to see if the active file path is already posted (in order to prevent duplicate postings). If required,
    proceeds to post the current file path as a comment for field reference.

    :param smartsheet.Smartsheet ss_c: Smartsheet client object
    :param int row_id: Smartsheet unique row identifying number
    :param str fpath: location of the quote in the local drive
    :return: Smartsheet server JSON response document
    :rtype: dict
    """
    # build comment object with filepath for text
    new_comment: smartsheet.Smartsheet.models.Comment = ss_c.models.Comment({'text': fpath})

    # grab existing discussions (if they exist) to help decide what to update
    current_discussions = ss_c.Discussions.get_row_discussions(sheet_id, row_id, include_all=True)

    # loop through discussions on row to identify if any comments already contain the current file path, if so, update
    discussion: smartsheet.Smartsheet.models.Discussion
    for discussion in current_discussions.data:
        for comment in discussion.comments:
            if comment.text == fpath:
                return ss_c.Discussions.update_comment(sheet_id, comment.id, new_comment)

    # add/update filepath to the project folder as a row comment
    return ss_c.Discussions.create_discussion_on_row(
        sheet_id, row_id, ss_c.models.Discussion({'comment': new_comment}))


def upload_row_info(ss_c, wb, row_id):
    """
    Builds new Row object, loops through the col_opts global variable, and builds Cell objects with information provided
    in the Excel project file. Handles different data types once extracted from Excel as required.

    :param smartsheet.Smartsheet ss_c: Smartsheet client object
    :param xw.Book wb: Excel workbook from which to extract data
    :param row_id: Smartsheet unique row identifying number
    :return: Smartsheet server JSON response document
    :rtype: dict
    """
    # build SS row object and append cells for info push
    job_row: smartsheet.Smartsheet.models.Row = ss_c.models.Row()
    job_row.id = row_id

    # loop through all columns to construct a cell to append to row if information is provided
    col: str
    xl_cell: str
    for col, xl_cell in col_opts.items():
        # create Cell instance, assign customer pricing from quote, add cell to Row object
        if col == 'Customer Pricing':
            job_row.cells.append(
                create_cell(ss_c, col, wb=wb, sheet='QUOTE', xl_cell=xl_cell, is_float=True))
        # create SS cell instance, assign quote date (always today), add cell to Row
        elif col == 'Quote Date':
            job_row.cells.append(
                create_cell(ss_c, col, cell_val=datetime.date.today().strftime('%Y-%m-%d')))
        # create SS cell instance, assign deal total from quote, add cell to Row
        elif col == 'Quote Value':
            job_row.cells.append(
                create_cell(ss_c, col, wb=wb, sheet='QUOTE', xl_cell=xl_cell, is_float=True))
        # create Cell instance, assign boolean value, add cell to Row
        elif col == 'Special Case 1':
            job_row.cells.append(
                create_cell(ss_c, col, wb=wb, sheet='SCHEDULE', xl_cell=xl_cell, is_bool=True))
        # create Cell instance, assign quote status to Completed, add cell to Row
        elif col == 'Status':
            job_row.cells.append(
                create_cell(ss_c, col, cell_val='Completed'))
        # all other columns follow the same pattern if a value exists in the Excel file, append with basic Cell instance
        elif type(cell_val := wb.sheets['SCHEDULE'].range(xl_cell).value) is str:
            job_row.cells.append(
                create_cell(ss_c, col, cell_val=cell_val))

    return ss_c.Sheets.update_rows(sheet_id, [job_row])


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
    if type(so_no) in (None, float):
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
