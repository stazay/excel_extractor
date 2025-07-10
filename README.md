
# excel_extractor - a basic 'relational database' tool

The purpose of this code is to consolidate data from Excel spreadsheets.

It can link between these spreadsheets to extract data where there is a matching reference point.

Finally, Python returns a nested list (`db`), where each sub-list represents a row from the spreadsheet, which can be written back to an Excel workbook for later use.

---

# Installation

1. Clone this repository into your workspace and install `excel_extractor` using pip:

```bash
git clone https://github.com/stazay/excel_extractor.git
pip install .
```

2. Import the package in your Python code:

```python
import excel_extractor
```

---

# Description of Key Functions

- `define_backups()`  
  Defines backups for use with `extend_db_entries()`, in case the main search fails.

- `extract_datum()`  
  Searches for corresponding data within the queried cell and returns it.

- `create_db_entries()`  
  Creates the base nested list; all entries within the specified range are extracted into this list.

- `extend_db_entries()`  
  Adds data to existing entries in the nested list by matching queried data in a specified column of the workbook. Extracts corresponding data from desired columns and appends it to the entries.

- `write_db_to_excel_workbook()`  
  Writes your nested list (`db`) back into a Microsoft Excel workbook.

---

# Example Usage

---

### 1. Define the output nested list and output workbook

```python
import xlsxwriter as xlsx

db = []
output_workbook = xlsx.Workbook("output.xlsx")
output_worksheet = output_workbook.add_worksheet()
```

---

### 2. Define workbooks to extract data from

```python
import xlwings as xw

workbook_1 = xw.Book(r"C:\Users\James\Documents\workbook_1.xlsx")
workbook_2 = xw.Book(r"C:\Users\James\Documents\workbook_2.xlsx")
workbook_3 = xw.Book(r"C:\Users\James\Documents\workbook_3.xlsx")
```

---

### 3. Extract relevant data from workbook_1, sheet_index 0, columns A, B, C, F, M, Q

```python
create_db_entries(
    db=db,
    workbook=workbook_1,
    sheet_index=0,
    desired_columns=["A", "B", "C", "F", "M", "Q"],
    queried_rows=(1, 250),
    clean_datetime="%d/%m/%Y",
    print_statements=True
)
```

- Extracts columns A, B, C, F, M, and Q from rows 1 to 250 in sheet 0 of `workbook_1`.
- Dates are cleaned to `dd/mm/yyyy` format.
- Prints a statement after each entry is extracted.

---

### 4. Extend `db` with data from workbook_2, sheet_index 2, columns A and F

```python
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
```

- Searches column D in `workbook_2` sheet 2 for matches with `db[i][3]`.
- If found, appends data from columns A and F to `db`.
- No backups used.
- Datetime cleaning disabled.
- Checks previous entries to avoid redundant extraction.
- Prints a statement after each entry is extracted.

---

### 5. Extract data from workbook_3 with backup sheets

```python
workbook_3_backups = define_backups(
    workbook=workbook_3,
    sheet_index=8,
    desired_columns=["F", "H"],
    queried_index=3,
    queried_column="D"
)

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
```

- Defines a backup search on `workbook_3` sheet 8, columns F and H, searched by column D.
- The main search is on sheet 1, columns F and M, searched by column C.
- Uses the backup if no match found in the main search.
- Cleans datetime values.
- Checks previous entries to avoid redundant extraction.
- Prints a statement after each entry is extracted.

---

### 6. Write the nested list `db` to the output workbook

```python
write_db_to_excel_workbook(
    db=db,
    workbook="output.xlsx",
    print_statements=True
)
```

- Writes the consolidated data to `output.xlsx`.
- Prints a statement after each row is written.

---

# Notes

- Ensure your Excel files are closed or opened in read-only mode when accessed by `xlwings`.
- Test the tool on small datasets before large-scale use.
- Combine with `pandas` for more advanced data manipulation if needed.

---

# Contribution

Contributions and issues are welcome. Feel free to open an issue or pull request!

---

Happy data extracting!
