# Example-Automation-Project

### Summary
This project is intended to showcase example code modules that are designed to represent an estimating design process
where all components were originally done by hand, but were largely automated using Python. 

All modules were coded in Python using the following dependencies: xlwings, pandas, pdfrw, pypdf2, and 
smartsheet-python-sdk. 

Due to the use of xlwings, extensive interaction with Excel spreadsheets was required, so many
functions end with an xlwings API call.

### Structure
The structure of the project is shown below where the 5 primary python files that are called via xlwings are in the
parent scripts directory and all of their helper modules are contained in similarly named folders.
```
📂 scripts
┣ 📝 customer_files.py
┣ 📝 product_quote.py
┣ 📝 product_submittal.py
┣ 📝 sales_order.py
┣ 📝 smartsheet_update.py
┣ 📂 quote
┃ ┣ components.py
┃ ┣ connections.py
┃ ┣ controls.py
┃ ┣ descriptions.py
┃ ┣ f_component.py
┃ ┣ io.py
┃ ┣ kit_sizes.py
┃ ┗ kit_types.py
┣ 📂 salesorder
┃ ┣ assign.py
┃ ┗ extract.py
┣ 📂 smartsheet_utils
┃ ┣ create_objects.py
┃ ┣ dump_columns.py
┃ ┗ upload.py
┣ 📂 submittal
┃ ┣ DWG.py
┃ ┗ spec.py
┣ 📂 utils
┃ ┣ list_dwgs.py
┃ ┗ rename.py
```

### Disclaimer
These scripts have been simplified and re-coded to showcase coding structure and methodology. Any references to 
products or internal processes have been removed to protect the identity of the company.

### Description
The estimation cycle effectively uses a combination of three documents: a schedule containing engineering information, a 
quote indicating pricing to the customer, and a combined submittal package containing product-specific information. The
second and third documents are based on the manually-entered information contained in the engineered schedule.

The automation of these three documents required constructing what was effectively several ETL processes between Excel
documents and other Excel documents, Smartsheet, and our enterprise SQL databases. Data is entered by hand into the
engineered schedule spreadsheet from plan drawing packages, but the quote, submittal package, and data pipelines to 
Smartsheet and the SQL server were coded in Python and activated by ActiveX buttons in Excel via the xlwings package.

All unit testing was performed using specially designed Excel files due to the required interaction with Excel via 
xlwings; thus, no formal unit tests are contained within the code.

### Results
Automating these processes resulted in a roughly 60% efficiency increase so that we were able to more than double job 
estimates handled without increasing the number of employees. Additionally, it resulted in a drastic decrease in errors 
in both the estimating and order entry departments, thus creating fewer production/shipping errors to customers.
