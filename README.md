# excel-2-python

Data Flow Diagram: https://miro.com/app/board/uXjVKaiToC4=/

### Setup

- Requires Python 3.11
- Set up a Python virtual environment using `pip install virtualenv` then `virtualenv -p python3.11 venv`
- Activate virtual environment with `source venv/bin/activate`
- Install requirements with `pip install -r requirements.txt`
- Install ipykernel to run jupyter notebooks `pip install ipykernel`
- For AWS CLI follow https://docs.aws.amazon.com/cli/latest/userguide/sso-configure-profile-token.html ensure you use https://d-9c677c26a6.awsapps.com/start#/

### TODO:

1. FormulaGenerator
2. SeriesDependenciesBuilder
3. DAGSorter

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

### Bugs

- Cells that start with "'=" don't work
- Formulas that have strings with colons Eg. =CONCATENATE(A2,":",B2) don't work
- When a table is horizontal (header_location="left") then the value at the top of this table (the final value when turned sideways) must not be a text value
