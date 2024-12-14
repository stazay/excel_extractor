# excel_extractor - a basic 'relational database' tool

The purpose of this code is to be able to consolidate data from Excel spreadsheets. 

A further function allows it to use links between these spreadsheets - and extract data from other spreadsheets where there is a matching point of reference.

Finally, Python can then return a nested list (db) - each sub-list representing what would be a row from the spreadsheet - which can be written to an Excel workbook for later use.


# Installation

1. Start by copying this repository into your workspace. After which, you can install excel_extractor using pip.
````
git clone https://github.com/stazay/excel_extractor.git
pip install .
````

2. In your code, simply enter the following, and you'll be able to use this tool in your own code.
````
import excel_extractor
````

# Description of key functions included
* `define_backups()` - THIS FUNCTION IS USED TO DEFINE BACKUPS ARGUMENT, IF REQUIRED, FOR USE WITH EXTEND_DB_ENTRIES().

* `extract_datum()` - THIS FUNCTION SEARCHES FOR CORRESPONDING DATA WITHIN THE QUERIED CELL, AND RETURNS IT AS AN OUTPUT.

* `create_db_entries()` - THIS FUNCTION IS USED TO CREATE THE BASIS OF YOUR NESTED LIST; ALL ENTRIES WITHIN THE RANGE ARE WRITTEN TO THE NESTED LIST.

* `extend_db_entries()` - THIS FUNCTION IS USED TO ADD DATA TO EXISTING ENTRIES IN YOUR NESTED LIST; IT SEARCHES FOR MATCHES OF QUERIED DATA WITHIN THE QUERIED COLUMN OF THE WORKBOOK. IT THEN EXTRACTS CORRESPONDING DATA FROM THE DESIRED COLUMNS AND APPENDS IT TO ENTRIES WITHIN THE NESTED LIST.

* `write_db_to_excel_workbook()` - THIS FUNCTION IS USED TO TRANSFER YOUR DATABASE BACK INTO MICROSOFT EXCEL.


# Example of code being used
1. Start by defining the output nested list, and output workbook (if wanting to transport the data back to a workbook).
````
db = []
output_workbook = xlsx.Workbook("output.xlsx")
output_worksheet = output_workbook.add_worksheet()
````

____

2. Define any workbooks of interest, from which you wish to extract data from.
````
workbook_1 = xw.Book("C:\Users\James\Documents\workbook_1.xlsx")
workbook_2 = xw.Book("C:\Users\James\Documents\workbook_2.xlsx")
workbook_3 = xw.Book("C:\Users\James\Documents\workbook_3.xlsx")
````

____

3. Extract relevant data from workbook_1, sheet_index: 0, from columns "A", "B", "C", "F", "M" & "Q" to nested list (db).
````
create_db_entries(
db=db,
workbook=workbook_1,
sheet_index=0,
desired_columns=["A", "B", "C", "F", "M", "Q"],
queried_rows=(1, 250),
clean_datetime="%d/%m/%Y",
print_statements=True
)
````
* The above code defines that we are using workbook_1 to extract info from sheet_index:0.
* Next, all data from columns A, B, C, F, M and Q are being extracted to the nested list, within all rows between the range of 1 and 250.
* All datetime objects are being cleaned to the "dd/mm/yyyy" format (see datetime).
* After each entry extracted, a print statement will be made.
 
____
 
 4. Extract relevant data from workbook_2, sheet_index: "2", specifically from queried_columns: "A" and "F". 
````
extend_db_entries(
db=db,
workbook=workbook_2,
sheet_index=2,
desired_columns=["A","F"],
queried_index=3,
queried_column="D",
backups=[],
clean_datetime=False,
check_previous=True,
print_statements=True
)
````

The above code extracts additional data to the nested list from workbook_2.
* First, sheet_index:2 of workbook_2 is searched for matches of the `db[i][queried_index]` (3rd index in the item in the nested list and column F from workbook_1). 
* If a match is found in the entirety of column D, corresponding data from columns A and F are appended to the nested list entry.
* No backup searches are defined. 
* Datetime objects are not cleaned.
* Previous entry in the nested list is being checked for a match in the `db[i][queried_index]`, to save time and extract the same information extracted as previously.
* After each entry extracted, a print statement will be made.

____

5. Next, extract relevant data from workbook_3, sheet_index: "2", queried_column: "D". If there is no match found in previous search, then instead extract relevant data from workbook_3, sheet_index: "3" via queried_column: "C".
````
workbook_3_backups = define_backups(workbook_3, 8, ["F", "H"], 3, "D")

extend_db_entries(
db=db,
workbook=workbook_3,
sheet_index=1,
desired_columns=["F", "M"],
queried_index=3,
queried_column="C",
backups=[workbook_3_backups],
clean_datetime="%d/%m/%Y",
check_previous=True,
print_statements=True
)
````

The above code extracts additional data to the nested list from workbook_3, sheetindex_2. In case of not finding any match, workbook_3, sheet_index_8 is queried instead.

(a) 
* A `backup` is defined for the next stage of data extraction (from workbook_3) - NOTE: multiple backups can be defined.
* This is using workbook_3, sheet_index:8, extracting information from columns F and H, by searching for a match in the `db[i][queried_index]` (3rd index in the item in the nested list and column F from workbook_1).
* If a match is found in the entirety of column D, corresponding data from columns F and H are appended to the nested list entry.
* NOTE: The `backup` code is odby used if the next stage of the data extraction fails to find any match.

(b) 
* Finally, sheet_index:1 of workbook_3 is searched for matches of the `db[i][queried_index]` (3rd index in the item in the nested list and column F from workbook_1).
* If a match is found in the entirety of column C, corresponding data from columns F and M are appended to the nested list entry.
* Backup entries are defined in the section above (b).
* All datetime objects are being cleaned to the "dd/mm/yyyy" format (see datetime).
* Previous entry in the nested list is being checked for a match in the `db[i][queried_index]`, to save time and extract the same information extracted as previously.
* After each entry extracted, a print statement will be made.

____

6. Paste data from output nested list (db) to output workbook.
````
write_db_to_excel_workbook(
db=db,
workbook="output.xlsx",
print_statements=True
)
````

* Finally, this code will write the created output (defined `db`) into an excel workbook (defined `output.xlsx`).
* After each entry extracted, a print statement will be made.
