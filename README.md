# excel_extractor - Extract and merge Excel spreadsheets

The purpose of this code is to consolidate data from Excel spreadsheets. 

A further function allows it to use links between these spreadsheets - and extract data from other spreadsheets where there is a matching point of reference.

Finally, Python can return a nested list (`db`) — each sub-list representing what would be a row from the spreadsheet — which can be written to an Excel workbook for later use.

## Requirements

- Python 3.7+
- `xlsxwriter` package
- `xlwings` package

You can install the requirements via pip:

```bash
pip install xlsxwriter xlwings
```

## Installation

1. Clone the repository into your workspace:

```bash
git clone https://github.com/stazay/excel_extractor.git
```

2. Install the package using pip:

```bash
pip install .
```

3. Import the package in your code:

```python
import excel_extractor
```

## Description of key functions included

- `define_backups()` - Defines backups argument for use with `extend_db_entries()`.
- `extract_datum()` - Extracts data from a specified cell.
- `create_db_entries()` - Creates the initial nested list (`db`) from the specified Excel range.
- `extend_db_entries()` - Extends the nested list by searching for matches in other workbooks/sheets and appending corresponding data.
- `write_db_to_excel_workbook()` - Writes the nested list data back into an Excel workbook.

## Example Usage

```python
import xlsxwriter as xlsx
import xlwings as xw
from excel_extractor import create_db_entries, extend_db_entries, define_backups, write_db_to_excel_workbook

# 1. Define the nested list and output workbook
db = []
output_workbook = xlsx.Workbook("output.xlsx")
output_worksheet = output_workbook.add_worksheet()

# 2. Open workbooks of interest
workbook_1 = xw.Book("C:\Users\James\Documents\workbook_1.xlsx")
workbook_2 = xw.Book("C:\Users\James\Documents\workbook_2.xlsx")
workbook_3 = xw.Book("C:\Users\James\Documents\workbook_3.xlsx")

# 3. Extract data from workbook_1
create_db_entries(
    db=db,
    workbook=workbook_1,
    sheet_index=0,
    desired_columns=["A", "B", "C", "F", "M", "Q"],
    queried_rows=(1, 250),
    clean_datetime="%d/%m/%Y",
    print_statements=True
)

# 4. Extend db with data from workbook_2
extend_db_entries(
    db=db,
    workbook=workbook_2,
    sheet_index=2,
    desired_columns=["A", "F"],
    queried_index=3,
    queried_column="D",
    backups=[],
    clean_datetime=False,
    check_previous=True,
    print_statements=True
)

# 5. Define backups for workbook_3
workbook_3_backups = define_backups(workbook_3, 8, ["F", "H"], 3, "D")

# Extend db with data from workbook_3 using backups
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

# 6. Write nested list back to Excel workbook
write_db_to_excel_workbook(
    db=db,
    workbook="output.xlsx",
    print_statements=True
)
```
