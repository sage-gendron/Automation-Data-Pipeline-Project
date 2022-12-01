# scripts/quote/io.py
"""
author: Sage Gendron
Contains functions to interact with the quote sheet within the project Excel file.
Primarily assigns quoted packages to standard Excel cell ranges (where lookups pull pricing from a database), but also
alters quoted_by cell and can clear the quote sheet if necessary (on revisions).
"""
import xlwings as xw

# Excel Quote cell ranges (dictionaries) for packages
quote_kitqty: dict[str, str] = {
    'A': 'E15', 'B': 'E30', 'C': 'E45', 'D': 'E60', 'E': 'E75', 'F': 'E90', 'G': 'E105', 'H': 'E120', 'I': 'E135',
    'J': 'E150', 'K': 'E165', 'L': 'E180', 'M': 'E195', 'N': 'E210', 'O': 'E225', 'P': 'E240', 'Q': 'E255', 'R': 'E270',
    'S': 'E285', 'T': 'E300', 'U': 'E315', 'V': 'E330', 'W': 'E345', 'X': 'E360', 'Y': 'E375', 'Z': 'E390',
    'AA': 'E405', 'AB': 'E420', 'AC': 'E435', 'AD': 'E450', 'AE': 'E465', 'AF': 'E480', 'AG': 'E495', 'AH': 'E510',
    'AI': 'E525', 'AJ': 'E540', 'AK': 'E555', 'AL': 'E570', 'AM': 'E585', 'AN': 'E600'}
quote_kit_pnrange: dict[str, str] = {
    'A': 'F16:F29', 'B': 'F31:F44', 'C': 'F46:F59', 'D': 'F61:F74', 'E': 'F76:F89', 'F': 'F91:F104', 'G': 'F106:F119',
    'H': 'F121:F134', 'I': 'F136:F149', 'J': 'F151:F164', 'K': 'F166:F179', 'L': 'F181:F194', 'M': 'F196:F209',
    'N': 'F211:F224', 'O': 'F226:F239', 'P': 'F241:F254', 'Q': 'F256:F269', 'R': 'F271:F284', 'S': 'F286:F299',
    'T': 'F301:F314', 'U': 'F316:F329', 'V': 'F331:F344', 'W': 'F346:F359', 'X': 'F361:F374', 'Y': 'F376:F389',
    'Z': 'F391:F404', 'AA': 'F406:F419', 'AB': 'F421:F434', 'AC': 'F436:F449', 'AD': 'F451:F464', 'AE': 'F466:F479',
    'AF': 'F481:F494', 'AG': 'F496:F509', 'AH': 'F511:F524', 'AI': 'F526:F539', 'AJ': 'F541:F554', 'AK': 'F556:F569',
    'AL': 'F571:F584', 'AM': 'F586:F599', 'AN': 'F601:F614'}
quote_kit_pnqty: dict[str, str] = {
    'A': 'G16:G29', 'B': 'G31:G44', 'C': 'G46:G59', 'D': 'G61:G74', 'E': 'G76:G89', 'F': 'G91:G104', 'G': 'G106:G119',
    'H': 'G121:G134', 'I': 'G136:G149', 'J': 'G151:G164', 'K': 'G166:G179', 'L': 'G181:G194', 'M': 'G196:G209',
    'N': 'G211:G224', 'O': 'G226:G239', 'P': 'G241:G254', 'Q': 'G256:G269', 'R': 'G271:G284', 'S': 'G286:G299',
    'T': 'G301:G314', 'U': 'G316:G329', 'V': 'G331:G344', 'W': 'G346:G359', 'X': 'G361:G374', 'Y': 'G376:G389',
    'Z': 'G391:G404', 'AA': 'G406:G419', 'AB': 'G421:G434', 'AC': 'G436:G449', 'AD': 'G451:G464', 'AE': 'G466:G479',
    'AF': 'G481:G494', 'AG': 'G496:G509', 'AH': 'G511:G524', 'AI': 'G526:G539', 'AJ': 'G541:G554', 'AK': 'G556:G569',
    'AL': 'G571:G584', 'AM': 'G586:G599', 'AN': 'G601:G614'}
# Excel Quote 'Quoted By'. To be updated (along with fx quoted_by()) when quoting personnel changes.
quote_author: str = 'AA5'
sg_cell: str = 'AB7'
user1_cell: str = 'AB4'
user2_cell: str = 'AB8'
user3_cell: str = 'AB6'
# variables for clear_quote()
quote_template_loc: str = r'C:\Estimating\Customer\Project Template.xlsm'
text_quote_range: str = 'E14:G682'
quote_fx_range: str = 'H14:L682'


def assign_pn_to_quote(wb, sub_dwg_dict, package_quantities):
    """
    Iterates through packages and package quantities and assigns them to correct package locations on quote.

    :param xw.Book wb: xlwings Book representing combination schedule/quote file with sheet[1] being the quote
    :param dict sub_dwg_dict: dictionary of pkg_key: [dwg, pns, qtys] for easy lookup and placement into Excel
    :param dict package_quantities: dictionary of [package keys: quantities] for package totals to be input into quote
    :return: None - changes made to Excel file
    """
    # assign package quantities to quote package value
    k: str
    v: int
    for k, v in package_quantities.items():
        wb.sheets['QUOTE'].range(quote_kitqty[k]).value = v

    # assign components and component quantities to quote
    m: str
    n: list[str, list[str], list[int]]
    for m, n in sub_dwg_dict.items():
        # ensure component list length less than package size allowable on quote (14 rows)
        if len(n[1]) > 13:
            raise Exception(f"Package {m} is longer than the allowable package size on the quote.")
        wb.sheets['QUOTE'].range(quote_kit_pnrange[m]).options(transpose=True).value = n[1]
        wb.sheets['QUOTE'].range(quote_kit_pnqty[m]).options(transpose=True).value = n[2]


def quoted_by(wb):
    """
    Changes 'quoted by' signature and initials within quote according to file name initials

    :param xw.Book wb: Book representing project file
    :return: None - changes made to Excel file
    """
    quote = wb.sheets['QUOTE']
    jn: str = wb.fullname.split('\\')[-1].split('_')[-1]
    if 'SG' in jn:
        quote.range(quote_author).value = quote.range(sg_cell).value
    elif 'USER1' in jn:
        quote.range(quote_author).value = quote.range(user1_cell).value
    elif 'USER2' in jn:
        quote.range(quote_author).value = quote.range(user2_cell).value
    elif 'USER3' in jn:
        quote.range(quote_author).value = quote.range(user3_cell).value


def clear_quote(wb):
    """
    Copies blank quote sheet fields from template onto calling workbook. Only gets called if 'R.' is present in the
    filename indicating a revision.

    :param xw.Book wb: calling Book object
    :return: None - clears quote sheet fields for a fresh start
    """
    # create a Book object for the quote template file for blank format copy
    wb_template: xw.Book = xw.Book(quote_template_loc)
    quote_template = wb_template.sheets['QUOTE']
    quote_current = wb.sheets['QUOTE']

    # un-filter calling project file's quote sheet so cells are copied to the correct ranges
    quote_current.api.AutoFilter.ShowAllData()

    # copy text range from template to calling project
    quote_template.range(text_quote_range).copy(quote_current.range(text_quote_range))

    # copy formula range from template to calling project
    fx1 = quote_template.range(quote_fx_range).formula
    quote_current.range(quote_fx_range).formula = fx1

    # save live project file and close the template
    wb.save()
    wb_template.close()
