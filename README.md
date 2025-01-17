# excel-2-python

Converts Excel formulas to Python code. Similar to the open-source library 'formulas' but instead of providing formulas and cells, this is built specifically for Excel tables. Excel tables are converted to pandas dataframes and the formulas are converted to pandas operations.

Data Flow Diagram: https://miro.com/app/board/uXjVKaiToC4=/

### Setup

- Requires Python 3.11
- Set up a Python virtual environment using `pip install virtualenv` then `virtualenv -p python3.11 venv`
- Activate virtual environment with `source venv/bin/activate`
- Install requirements with `pip install -r requirements.txt`

### Requirements for reduced Excel

- Each column (vertical or horizontal) must have one header (must be characters and no pipes)
- No Excel "table" objects
- No pivot tables
- No hidden tabs
- No fixed columns in formulas
- All formulas are draggable
- No spaces or pipes in tab names
- No circular references
- No cell references to column headers in formulas
- No references to empty cells
- There must be exactly two rows for each column with formulas
- There must be two or more rows for each column with values
- Any column references must refer to a column with only one series of data with column header as first row
- Dates must be in number (general) format: Eg. 44501 or 45085.5
- Formulas in a horizontal series can only reference other cells from a horizontal series (same applies for vertical)

Supported functions:

- SUM
- AVERAGE
- IF
- INDEX
- MATCH
- COUNT
- COUNTIF
- COUNTA
- SUMIF
- REPLACE
- CONCATENATE
- LEFT
- RIGHT
- LEN
- TRIM
- LOWER
- UPPER
- ROUND
- AND
- OR
- MAX
- MIN
- DATE
- DAY
- ROUNDDOWN
- ROUNDUP
- TRIM
