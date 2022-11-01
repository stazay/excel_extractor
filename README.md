# excel_extractor - a basic 'relational database' tool

The purpose of this code is to be able to consolidate data from Excel spreadsheets. 

A further function allows it to use links between these spreadsheets - and extract data from other spreadsheets where there is a matching point of reference.

Finally, Python can then return a dataframe - which can be used with Pandas, or be written to an Excel workbook for later use.


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
define_backups() - THIS FUNCTION IS USED TO DEFINE BACKUPS ARGUMENT, IF REQUIRED, FOR USE WITH EXTEND_DF_ENTRIES().
````
define_backups(workbook, sheet_index, desired_columns, queried_index, queried_column)
        - workbook          -- the workbook being queried
        - sheet_index       -- the worksheet being queried
        - desired_columns   -- a list containing all of the columns from which to extract data from
                                eg: ['A', 'B', 'F']
                                will return data entries from columns A, B and F in Excel
        - queried_index     -- the input being queried for matches within the queried_column;
                                namely an index of the data in a dataframe entry
                                (returning the queried_datum)
                                eg: 3
                                will take i[3] from the entry in the dataframe,
                                and search for a match within the queried_column
        - queried_column    -- the column of information that is being queried for a match against
                                the queried_datum
````
____

extract_datum() - THIS FUNCTION SEARCHES FOR CORRESPONDING DATA WITHIN THE QUERIED CELL, AND RETURNS IT AS AN OUTPUT.
````
extract_datum(workbook, sheet_index, queried_column, queried_row)
        - workbook          -- the workbook being queried
        - sheet_index       -- the worksheet being queried
        - queried_column    -- the column of information that data is being extracted from
        - queried_row       -- the row of information that data is being extracted from
````
____

create_df_entries() - THIS FUNCTION IS USED TO CREATE THE BASIS OF YOUR DATAFRAME; ALL ENTRIES WITHIN THE RANGE ARE WRITTEN TO THE DATAFRAME.
````
create_df_entries(df, workbook, sheet_index, desired_columns, queried_rows, clean_datetime, print_statements)
        - df                -- the dataframe being written
        - workbook          -- the workbook being queried
        - sheet_index       -- the worksheet being queried
        - desired_columns   -- a list containing all of the columns from which to extract data from
                                eg: ['A', 'B', 'F']
                                will return data entries from columns A, B and F in Excel
        - queried_rows      -- a tuple containing: the range of cells of interest
                                (if left empty, it will assume the that all rows are of interest)
                                eg: (0, 15)
                                will return data entries from rows 0-15
        - clean_datetime    -- will clean datetime objects into the desired format input as a
                                string, see datetime.strftime (False by default)
                                eg: '%d/%m/%Y' will return 'dd/mm/yyyy'
        - print_statements  -- will return print-statements outlining progress of data extraction
                                if True (True by default)
````
____

extend_df_entries() - THIS FUNCTION IS USED TO ADD DATA TO EXISTING ENTRIES IN YOUR DATAFRAME; IT SEARCHES FOR MATCHES OF QUERIED DATA WITHIN THE QUERIED COLUMN OF THE WORKBOOK. IT THEN EXTRACTS CORRESPONDING DATA FROM THE DESIRED COLUMNS AND APPENDS IT TO ENTRIES WITHIN THE DATAFRAME.
````
extend_df_entries(df, workbook, sheet_index, desired_columns, queried_index, queried_column, backups, 
                    clean_datetime, check_previous, print_statements)
        - df                -- the dataframe being written
        - workbook          -- the workbook being queried
        - sheet_index       -- the worksheet being queried
        - desired_columns   -- a list containing all of the columns from which to extract data from
                                eg: ['A', 'B', 'F']
                                will return data entries from columns A, B and F in Excel
        - queried_index     -- the input being queried for matches within the queried_column;
                                namely an index of the data in a dataframe entry
                                (returning the queried_datum)
                                eg: 3
                                will take i[3] from the entry in the dataframe,
                                and search for a match within the queried_column
        - queried_column    -- the column of information that is being queried for a match against
                                the queried_datum
        - backups           -- this is used as a contingency when no match is found from
                                queried_input within the queried_column - use to provide a secondary
                                extraction step
                                eg: [[backup_workbook, 3, ['A', 'B', 'F'], 3, 'C']]
                                ^ list containing the following:
                                backup_workbook, backup_sheet_index, desired_columns,
                                queried_index, queried_column
                                NOTE: it is possible to provide multiple backup lists within a
                                single list as the input
        - clean_datetime    -- will clean datetime objects into the desired format input as a
                                string, see datetime.strftime (False by default)
                                eg: '%d/%m/%Y'
                                will return 'dd/mm/yyyy'
        - check_previous    -- will check previous entry for any match with the queried_datum,
                                and copy previous information to save time (False by default)
        - print_statements  -- will return print-statements outlining progress of data extraction
                                if True (True by default)
````
____

write_df_to_excel_workbook() - THIS FUNCTION IS USED TO TRANSFER YOUR DATABASE BACK INTO MICROSOFT EXCEL.
````
write_df_to_excel_workbook(df, workbook, print_statements)
        - df                -- the dataframe being extracted from
        - workbook          -- the workbook being written
        - print_statements  -- will return print-statements outlining progress of data extraction
                                if True (True by default)
````


# Example of code being used

1. Start by defining the dataframe, and output workbook (if wanting to transport the data back to a workbook).
````
df = []
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

3. Extract relevant data from workbook_1, sheet_index: 0, from columns "A", "B", "C", "F", "M" & "Q" to dataframe (df).
````
create_df_entries(
df=df,
workbook=workbook_1,
sheet_index=0,
desired_columns=["A", "B", "C", "F", "M", "Q"],
queried_rows=(1, 250),
clean_datetime="%d/%m/%Y",
print_statements=True
)
````
The above code defines that we are using workbook_1 to extract info from sheet_index:0.
Next, all data from columns A, B, C, F, M and Q are being extracted to the dataframe, within all rows between the range of 1 and 250.
All datetime objects are being cleaned to the "dd/mm/yyyy" format (see datetime).
After each entry extracted, a print statement will be made.
 
____
 
 4. Extract relevant data from workbook_2, sheet_index: "2", specifically from queried_columns: "A" and "F". Next, extract relevant data from workbook_3, sheet_index: "2", queried_column: "D". If there is no match found in previous search, then instead extract relevant data from workbook_3, sheet_index: "3" via queried_column: "C".
````
extend_df_entries(
df=df,
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

workbook_3_backups = define_backups(workbook_3, 8, ["F", "H"], 3, "D")

extend_df_entries(
df=df,
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

The above code extracts additional data to the dataframe from workbook_2 and workbook_3.


(a) First, sheet_index:2 of workbook_2 is searched for matches of the `df[i][queried_index]` (3rd index in the item in the dataframe and column F from workbook_1). 
If a match is found in the entirety of column D, corresponding data from columns A and F are appended to the dataframe entry.
No backup searches are defined. 
Datetime objects are not cleaned.
Previous entry in the dataframe is being checked for a match in the `df[i][queried_index]`, to save time and extract the same information extracted as previously.
After each entry extracted, a print statement will be made.


(b) Next, a `backup` is defined for the next stage of data extraction (from workbook_3).
This is using workbook_3, sheet_index:8, extracting information from columns F and H, by searching for a match in the `df[i][queried_index]` (3rd index in the item in the dataframe and column F from workbook_1).
If a match is found in the entirety of column D, corresponding data from columns F and H are appended to the dataframe entry.
NOTE: The `backup` code is only used if the next stage of the data extraction fails to find any match.


(c) Finally, sheet_index:1 of workbook_3 is searched for matches of the `df[queried_index]` (3rd index in the item in the dataframe and column F from workbook_1).
If a match is found in the entirety of column C, corresponding data from columns F and M are appended to the dataframe entry.
Backup entries are defined in the section above (b).
All datetime objects are being cleaned to the "dd/mm/yyyy" format (see datetime).
Previous entry in the dataframe is being checked for a match in the `df[queried_index]`, to save time and extract the same information extracted as previously.
After each entry extracted, a print statement will be made.

____

5. Paste data from dataframe (df) to output workbook.
````
write_df_to_excel_workbook(
df=df,
workbook="output.xlsx",
print_statements=True
)
````

Finally, this code will write the created dataframe into an excel workbook (defined output.xlsx).
After each entry extracted, a print statement will be made.
