# scripts/smartsheet_utils/create_objects.py
"""
author: Sage Gendron
Helper modules to assist in uploading information to Smartsheet servers. As Smartsheet's API is a little tricky, wrote
functions that are a little more straightforward for common use.
"""
import datetime
import smartsheet

from create_objects import create_cell

# unique ID for estimating Smartsheet sheet
sheet_id: str = '@@@'

# schedule Excel file cell variables for update_smartsheet()
col_opts: dict[str, str] = {
    'City': 'D8', 'Customer Pricing': 'A3', 'Engineer': 'C22', 'Engineered Component': 'D4', 'Equipment Type': 'D6',
    'Industry': 'D3', 'Phase': 'G24', 'Quote Date': None, 'Quote Value': 'AA4', 'Special Case 1': 'D5', 'State': 'D9',
    'Status': None, 'Street Address': 'D7',  'Zip': 'D10'
}


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
