"""
excel_extractor - a basic relational database creator       - Saba Tazayoni, 15/10/2021

The purpose of this code is to be able to consolidate data from Excel spreadsheets. 
A further function allows it to then form a relational database - and extract data from other spreadsheets 
    where there is a matching point of reference.
Finally, Python can then return a dataframe - which can be used with Pandas, or be written to an Excel workbook for later use.
"""

from datetime import datetime
import sys
import xlsxwriter as xlsx
import xlwings as xw


# LOOSE FUNCTIONS
# LOOSE FUNCTIONS
def check_total_rows(workbook, sheet_index):
    ### THIS FUNCTION CHECKS NUMBER OF ROWS IN THE WORKSHEET.
    # workbook          -- the workbook being queried
    # sheet_index       -- the worksheet being queried
    worksheet = workbook.sheets[sheet_index]
    last_row = worksheet.range('A' + str(worksheet.cells.last_cell.row)).end('up').row
    
    return last_row


def check_row_number(workbook, sheet_index, queried_column, queried_datum):
    ### THIS FUNCTION LOOKS FOR CORRESPONDING DATA WITHIN THE QUERIED COLUMN, AND RETURNS A ROW NUMBER WITH A MATCH.
    # workbook          -- the workbook being queried
    # sheet_index       -- the worksheet being queried
    # queried_column    -- the column of information that data is being extracted from
    # queried_datum      -- an input that corresponds to wanted info from the queried_column
    last_row = check_total_rows(workbook, sheet_index)

    try:
        for row_number in range(1, last_row):
            queried_cell = workbook.sheets[sheet_index].range(f"{queried_column}{row_number}").value
            if (queried_datum == queried_cell):
                return row_number
    except:
        return None


def clean_datetime_object(input, format):
    # THIS FUNCTION CLEANS DATETIME OBJECTS INTO A DESIRED FORMAT
    # input             -- the object being amended
    # format            -- the desired format, see datetime.strftime
    #                       eg: "%d/%m/%Y" will return "dd/mm/yyyy"
    if (type(input) == datetime):
        try:
            output = input.strftime(f"{format}")
            return output
        except:
            return input
    
    return input


def define_backups(workbook, sheet_index, desired_columns, queried_df_index, queried_column):
    ### THIS FUNCTION IS USED TO DEFINE BACKUPS ARGUMENT, IF REQUIRED, FOR EXTRACT_ADDITIONAL_DATA_TO_DF()
    # workbook          -- the workbook being queried
    # sheet_index       -- the worksheet being queried
    # desired_columns   -- a list containing all of the columns from which to extract data from
    #                       eg: ["A", "B", "F"]     will return data entries from columns A, B and F in Excel
    # queried_df_index  -- the input being queried for matches within the queried_column; namely an index of the data in a dataframe entry (returning the queried_datum)
    #                       eg: 3                   will take i[3] from the entry in the dataframe, and search for a match within the queried_column
    # queried_column    -- the column of information that is being queried for a match against the queried_datum
    try:
        backups = [workbook, sheet_index, desired_columns, queried_df_index, queried_column]
        return backups      
    except:
        print("Backups not configured in usable format.")
        sys.exit(1)


def extract_datum(workbook, sheet_index, queried_column, queried_row):
    ### THIS FUNCTION SEARCHES FOR CORRESPONDING DATA WITHIN THE QUERIED CELL, AND RETURNS IT AS AN OUTPUT.
    # workbook          -- the workbook being queried
    # sheet_index       -- the worksheet being queried
    # queried_column    -- the column of information that data is being extracted from
    # queried_row       -- the row of information that data is being extracted from
    output = workbook.sheets[sheet_index].range(f"{queried_column}{queried_row}").value

    return output


# LARGER FUNCTIONS
# LARGER FUNCTIONS
def extract_new_entries_to_df(df, workbook, sheet_index, desired_columns, queried_rows="default", clean_datetime=False, print_statements=True):
    ### THIS FUNCTION IS USED TO CREATE THE BASIS OF YOUR DATAFRAME; ALL ENTRIES WITHIN THE RANGE ARE WRITTEN TO THE DATAFRAME
    # df                -- the dataframe being written
    # workbook          -- the workbook being queried
    # sheet_index       -- the worksheet being queried
    # desired_columns   -- a list containing all of the columns from which to extract data from
    #                       eg: ["A", "B", "F"] will return data entries from columns A, B and F in Excel
    # queried_rows      -- a tuple containing: the range of cells of interest (if left empty, it will assume the that all rows are of interest)
    #                       eg: (0, 15) will return data entries from rows 0-15
    # clean_datetime    -- will clean datetime objects into the desired format input as a string, see datetime.strftime (False by default)
    #                       eg: "%d/%m/%Y" will return "dd/mm/yyyy"
    # print_statements  -- will return print-statements outlining progress of data extraction if True (True by default)
    if (queried_rows == "default"):
        first_row = 1
        last_row = check_total_rows(workbook, sheet_index)
    else:
        try:
            (first_row, last_row) = queried_rows
        except:
            print(f"Queried Range: {queried_rows} must be a TUPLE, containing only TWO numbers")
            sys.exit(1)
        
    try:
        for i in range(first_row, last_row):
            row_data = []
            for column in desired_columns:
                output_data = extract_datum(workbook, sheet_index, f"{column}", i)
                
                # clean the datetime objects
                if (clean_datetime != False):
                    output_data = clean_datetime_object(output_data, clean_datetime)

                row_data.append(output_data)
            df.append(row_data)

            # print progress
            if print_statements:
                progress = int((i / last_row) * 100)
                print(f"Extracting data from {workbook}: {progress}% -- {row_data}")

        if print_statements:
            print(f"Extracting data from {workbook}: 100%")
        return df

    except:
        print(f"Queried Range: {queried_rows} must be a TUPLE, containing only TWO numbers, whereby the second number is bigger than the first number")
        sys.exit(1)


def extract_additional_data_to_df(df, workbook, sheet_index, desired_columns, queried_df_index, queried_column, backups=[], clean_datetime=False, check_previous=False, print_statements=True):
    ### THIS FUNCTION IS USED TO ADD DATA TO EXISTING ENTRIES IN YOUR DATAFRAME; IT SEARCHES FOR MATCHES OF QUERIED DATA WITHIN THE QUERIED COLUMN OF THE WORKBOOK.
    ### IT THEN EXTRACTS CORRESPONDING DATA FROM THE DESIRED COLUMNS AND APPENDS IT TO THE DATAFRAME.
    # df                -- the dataframe being written
    # workbook          -- the workbook being queried
    # sheet_index       -- the worksheet being queried
    # desired_columns   -- a list containing all of the columns from which to extract data from
    #                       eg: ["A", "B", "F"]     will return data entries from columns A, B and F in Excel
    # queried_df_index  -- the input being queried for matches within the queried_column; namely an index of the data in a dataframe entry (returning the queried_datum)
    #                       eg: 3                   will take i[3] from the entry in the dataframe, and search for a match within the queried_column
    # queried_column    -- the column of information that is being queried for a match against the queried_datum
    # backups           -- this is used as a contingency when no match is found from queried_input within the queried_column - use this to provide a secondary extraction step
    #                       eg: [[backup_workbook, 3, ["A", "B", "F"], 3, "C"]]
    #                         ^ list containing the following: a backup_workbook, backup_sheet_index, desired_columns, queried_df_index, queried_column
    #                       NOTE: it is possible to provide multiple backup lists within a single list as the input
    # clean_datetime    -- will clean datetime objects into the desired format input as a string, see datetime.strftime (False by default)
    #                       eg: "%d/%m/%Y"          will return "dd/mm/yyyy"
    # check_previous    -- will check previous X entries for any match with the queried_datum, and copy previous information to save time (False by default)
    # print_statements  -- will return print-statements outlining progress of data extraction if True (True by default)
    for (index, i) in enumerate(df):
        queried_datum = i[queried_df_index]
        
        # check previous to save time
        if (check_previous == True) and (len(df[index-1]) > 0) and ((df[index-1][queried_df_index]) == queried_datum):
            entry_difference = (len(df[index-1]) - len(i))
            for j in range(entry_difference):
                i.append(None)
                i[(len(i)-1)] = df[index-1][(len(i)-1)]

        else:
            row_number = check_row_number(workbook, sheet_index, queried_column, queried_datum)

            # iterate through all desired_columns of data
            if (row_number != None):
                for column in desired_columns:
                    output_data = extract_datum(workbook, sheet_index, column, row_number)
                    
                    # clean the datetime objects
                    if (clean_datetime != False):
                        output_data = clean_datetime_object(output_data, clean_datetime)

                    i.append(output_data)
        
            # iterate through all desired_columns of data in backup inputs
            elif (backups != []):
                # iterate through all backups
                for backup in backups:
                    try:
                        bu_workbook = backup[0]
                        bu_worksheet = backup[1]
                        bu_desired_columns = backup[2]
                        bu_queried_index = backup[3]
                        bu_queried_column = backup[4]
                        bu_queried_datum = i[bu_queried_index]

                        row_number = check_row_number(bu_workbook, bu_worksheet, bu_queried_column, bu_queried_datum)

                        # iterate through all desired_columns of data
                        if row_number != None:
                            for column in bu_desired_columns:
                                output_data = extract_datum(bu_workbook, bu_worksheet, column, row_number)

                                # clean the datetime objects
                                if (clean_datetime != False):
                                    output_data = clean_datetime_object(output_data, clean_datetime)

                                i.append(output_data)
                            break  
                    except:
                        for column in desired_columns:
                            i.append(None)

            # if no matches, append blank spaces
            else: 
                for column in desired_columns:
                    i.append(None)

        if print_statements:
            progress = ((index/len(df))*100)
            print(f"Extracting data from {workbook}: {int(progress)}% -- {i}")
        
    print(f"Extracting data from {workbook}: 100%")
    return df


def write_df_to_excel_workbook(df, workbook, print_statements=True):
    ### THIS FUNCTION IS USED TO TRANSFER YOUR DATABASE BACK INTO MICROSOFT EXCEL
    # df                -- the dataframe being extracted from
    # workbook          -- the workbook being written
    # print_statements  -- will return print-statements outlining progress of data extraction if True (True by default)
    output_workbook = xlsx.Workbook(f"{workbook}")
    output_worksheet = output_workbook.add_worksheet()

    row = 0

    for (index, i) in enumerate(df):
        for (col, j) in enumerate(i):
            output_worksheet.write(row, col, j)
        row += 1

        # Display Progress
        if print_statements:
            progress = ((index/len(df))*100)
            print(f"Extracting data from dataframe: {int(progress)}%: {i}")

    if print_statements:
        print(f"Extracting data from {workbook}: 100%")
    output_workbook.close()


# IMPORTANT
# IMPORTANT
if __name__ == "__main__":
    pass
