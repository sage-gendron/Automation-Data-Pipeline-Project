# smartsheet_update.py
import datetime
import smartsheet
import xlwings as xw
from rename import rename
"""
author: Sage Gendron

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
    'Status': 111000, 'Estimator': 222000, 'Company': 444000, 'Opportunity': 555000, 'Customer Pricing': 666000,
    'Quote Value': 777000, 'Quote Date': 101010, 'Date Won': 121212, 'Deal Status': 131313, 'Industry': 161616,
    'Engineered Component': 171717, 'Special Case 1': 181818, 'Equipment Type': 191919, 'Engineer': 202020,
    'Street Address': 212121, 'City': 222222, 'State': 232323, 'Zip': 242424, 'SO #': 262626, 'Phase': 272727
}

# schedule Excel file cell variables for update_smartsheet()
simple_cols: dict[str, str] = {
    'Industry': 'D3', 'Engineered Component': 'D4', 'Special Case 1': 'D5', 'Equipment Type': 'D6', 'Engineer': 'C22',
    'Street Address': 'D7', 'City': 'D8', 'State': 'D9', 'Zip': 'D10', 'Phase': 'G24', 'Customer Pricing': 'A3',
    'Status': None
}
# for logical updating
sch_owner_cell: str = 'G23'
sch_row_id_cell: str = 'D2'
qte_total_cell: str = 'AA4'
so_cell: str = 'G22'


def get_ss_client(wb):
    """
    Uses filepath from the passed workbook to identify author and create a Smartsheet client using their API key.

    :param xw.Book wb: Excel workbook to grab the estimator's initials from
    :returns ss_c: A smartsheet client object created with the estimator's API key
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
    :param bool is_bool: indicates if values should be read as a bool
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

    new_cell.strict = False

    return new_cell


def update_ss():
    """
    Updates the Quotes Smartsheet with information from the calling project file as detailed by created cells in this
    function. Also checks for existing attachments and discussions to update in lieu of create (where applicable).

    :return: None
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()

    fpath = '\\'.join(wb.fullname.split('\\')[:-1])

    quote_name = rename(wb, 'QUOTE', 'pdf')
    opt_quote_name = rename(wb, 'EXCEL QUOTE', 'xlsx')
    schedule_name = rename(wb, 'SCHEDULE', 'xlsx')
    submittal_name = rename(wb, 'SUBMITTAL', 'pdf')

    # instantiate smartsheet client with Quotes smartsheet
    ss_c: smartsheet.Smartsheet = get_ss_client(wb)

    # build SS row object and append cells for info push
    job_row_id: int = int(wb.sheets['SCHEDULE'].range(sch_row_id_cell).value)
    job_row: smartsheet.Smartsheet.models.Row = ss_c.models.Row()
    job_row.id = job_row_id

    # build comment object with filepath for text
    fpath_comment: smartsheet.Smartsheet.models.Comment = ss_c.models.Comment({'text': fpath})

    # instantiate flags to skip update calls below
    update_fpath_comment = None
    update_quote_upload = None
    update_schedule_upload = None
    update_submittal_upload = None
    update_opt_quote_upload = None

    # grab existing attachments, discussions to help decide what to update
    current_attachments = ss_c.Attachments.list_row_attachments(sheet_id, job_row_id)
    current_discussions = ss_c.Discussions.get_row_discussions(sheet_id, job_row_id, include_all=True)

    discussion: smartsheet.Smartsheet.models.Discussion
    for discussion in current_discussions.data:
        for comment in discussion.comments:
            if comment.text == fpath:
                update_fpath_comment = comment.id

    attachment: smartsheet.Smartsheet.models.Attachment
    for attachment in current_attachments.data:
        if attachment.name == quote_name.split('\\')[-1]:
            update_quote_upload = attachment.id
        elif attachment.name == schedule_name.split('\\')[-1]:
            update_schedule_upload = attachment.id
        elif attachment.name == submittal_name.split('\\')[-1]:
            update_submittal_upload = attachment.id
        elif attachment.name == opt_quote_name.split('\\')[-1]:
            update_opt_quote_upload = attachment.id

    # create SS cell instance, assign deal total from quote, add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Quote Value', xl_cell=qte_total_cell, wb=wb, sheet='QUOTE', is_float=True))

    # create SS cell instance, assign quote date (always today), add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Date of Quote', cell_val=datetime.date.today().strftime('%Y-%m-%d')))

    # loop through all columns to construct a cell to append to row if information is provided
    col: str
    xl_cell: str
    for col, xl_cell in simple_cols.items():
        # create Cell instance, assign quote status to Completed, add cell to SS row object
        if col == 'Status':
            job_row.cells.append(create_cell(ss_c, col, cell_val='Completed'))
        # create Cell instance, assign boolean value, add cell to SS row object
        elif col == 'Special Case 1':
            job_row.cells.append(create_cell(ss_c, 'Special Case 1', wb=wb, sheet='SCHEDULE', xl_cell=xl_cell,
                                             is_bool=True))
        # create Cell instance, assign customer pricing from quote, add cell to SS row object
        elif col == 'Customer Pricing':
            job_row.cells.append(
                create_cell(ss_c, col, wb=wb, xl_cell=xl_cell, sheet='QUOTE', is_float=True))
        elif type(cell_val := wb.sheets['SCHEDULE'].range(xl_cell).value) is str:
            job_row.cells.append(create_cell(ss_c, col, cell_val=cell_val))

    # push row update to SS server
    _updated_row = ss_c.Sheets.update_rows(sheet_id, [job_row])

    # add/update filepath to the project folder as comment
    if update_fpath_comment is not None:
        _updated_comment = ss_c.Discussions.update_comment(sheet_id, update_fpath_comment, fpath_comment)
    else:
        _updated_comment = ss_c.Discussions.create_discussion_on_row(
            sheet_id, job_row_id, ss_c.models.Discussion({'comment': fpath_comment}))

    # initiate file attachment upload to row by filename/path
    if update_quote_upload is not None:
        _upload_quote = ss_c.Attachments.attach_new_version(sheet_id, update_quote_upload, (
            quote_name.split('\\')[-1], open(quote_name, 'rb'), 'application/pdf'))
    else:
        _upload_quote = ss_c.Attachments.attach_file_to_row(sheet_id, job_row_id, (quote_name.split('\\')[-1],
                                                            open(quote_name, 'rb'), 'application/pdf'))

    #
    try:
        if update_submittal_upload is not None:
            _upload_submittal = ss_c.Attachments.attach_new_version(sheet_id, update_submittal_upload, (
                submittal_name.split('\\')[-1], open(submittal_name, 'rb'), 'application/pdf'))
        else:
            _upload_submittal = ss_c.Attachments.attach_file_to_row(sheet_id, job_row_id, (
                submittal_name.split('\\')[-1], open(submittal_name, 'rb'), 'application/pdf'))
    except FileNotFoundError:
        pass

    #
    if update_schedule_upload is not None:
        _upload_schedule = ss_c.Attachments.attach_new_version(
            sheet_id, update_schedule_upload, (schedule_name.split('\\')[-1], open(schedule_name, 'rb'),
                                               'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
    else:
        _upload_schedule = ss_c.Attachments.attach_file_to_row(
            sheet_id, job_row_id, (schedule_name.split('\\')[-1], open(schedule_name, 'rb'),
                                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))

    #
    if update_opt_quote_upload is not None:
        try:
            _upload_opt_quote = ss_c.Attachments.attach_new_version(
                sheet_id, update_opt_quote_upload, (opt_quote_name.split('\\')[-1], open(opt_quote_name, 'rb'),
                                                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
        except FileNotFoundError:
            pass

    else:
        try:
            _upload_opt_quote = ss_c.Attachments.attach_file_to_row(
                sheet_id, job_row_id, (opt_quote_name.split('\\')[-1], open(opt_quote_name, 'rb'),
                                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
        except FileNotFoundError:
            pass


def mark_as_won():
    """
    Changes the deal status of the active quote's row ID to 'Won', changes the date won to today, and adds the SO#
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

    # create SS cell instance, assign quote status information, add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Status', cell_val='Won'))

    # create SS cell instance, assign date won (always today), add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'Date Won', cell_val=datetime.date.today().strftime('%Y-%m-%d')))

    # create SS cell instance, assign sales order # (NoneType if cell is blank), add cell to SS row object
    job_row.cells.append(create_cell(ss_c, 'SO #', cell_val=so_no))

    # push SS row object as an update to SS server via previously instantiated Sheets object
    _updated_row = ss_c.Sheets.update_rows(sheet_id, [job_row])
