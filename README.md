# Example-Automation-Project

### Summary
This is intended to showcase some example code modules that were designed to represent an estimating design process
where all components were originally done by hand, but have now been 

All modules were coded in Python using xlwings, pandas, pdfrw, pypdf2, and smartsheet-python-sdk.

### Disclaimer
This has been re-coded to indicate the gists of the processes and to protect the identity of the company, its internal
processes, and its products.

### Description
The automated cycle effectively used a combination of three documents: a schedule containing engineering information, a 
quote indicating pricing to the customer, and a combined submittal package containing product-specific information. The
second and third documents are based on the information contained in the engineered schedule.

The automation of these three documents required constructing what was effectively several ETL processes between Excel
documents and other Excel documents, Smartsheet, and our enterprise SQL databases. Data is entered by hand into the
engineered schedule spreadsheet from plan drawing packages, but the quote, submittal package, and data pipelines to 
Smartsheet and the SQL server were coded in Python and activated by ActiveX buttons in Excel via the xlwings package.

All unit testing was performed using example Excel files due to the required interaction with Excel/xlwings.