# Automation & Data Pipeline Project

### Summary
This project is intended to showcase example code modules that are designed to represent an estimating design process
where all components were originally done by hand, but were largely automated using Python. 

All modules are written in Python using the following dependencies: xlwings, pandas, pdfrw, pypdf2, and 
smartsheet-python-sdk. 

Due to the use of xlwings, extensive interaction with Excel spreadsheets was required, so many
functions end with an xlwings API call.

### Disclaimer
These scripts have been simplified and re-coded to showcase coding structure and methodology. Any references to 
products or internal processes have been removed to protect the identity of the company.

### Structure


### Description
The estimation cycle uses a combination of three documents: 

```
 - an engineered schedule containing estimate-specific data provided by engineers and estimate-specifc product 
   information
 - a quote indicating pricing to the customer and containing a bill of materials
 - a combined submittal package containing product-specific information. 
```

The quote and submittal documents are based on the information contained in the engineered schedule.

The automation of these three documents required constructing what was effectively several ETL processes between Excel
documents and other Excel documents, Smartsheet, and our enterprise SQL databases. Data is entered by hand into the
engineered schedule spreadsheet from plan drawing packages, but the quote, submittal package, and data pipelines to 
Smartsheet and the SQL server were coded in Python and activated by ActiveX buttons in Excel via the xlwings package.

All unit testing was performed using specially designed Excel files due to the required interaction with Excel via 
xlwings; thus, no formal unit tests are contained within the code.

### Results
Automating these processes resulted in a ~60% estimating efficiency increase so that the department was able to more than double 
job estimates handled without increasing the number of employees. Additionally, it resulted in a drastic decrease in errors 
in both the estimating and order entry departments, thus creating fewer production/shipping errors and increasing throughput.
