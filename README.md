# excel-2-python

Data Flow Diagram: https://miro.com/app/board/uXjVKaiToC4=/

### Setup

- Requires Python 3.11
- Set up a Python virtual environment using `pip install virtualenv` then `virtualenv -p python3.11 venv`
- Activate virtual environment with `source venv/bin/activate`
- Install requirements with `pip install -r requirements.txt`
- Install ipykernel to run jupyter notebooks `pip install ipykernel`
- For AWS CLI follow https://docs.aws.amazon.com/cli/latest/userguide/sso-configure-profile-token.html ensure you use https://d-9c677c26a6.awsapps.com/start#/

### Excel requirements

- Each column (vertical or horizontal) must have one or two headers (must be characters not numbers)
- No Excel "table" objects
- No pivot tables
- Tab names do not have spaces
- No hidden tabs
- No circular references
- Dates must be in number (general) format: Eg. 44501 or 45085.5
- Only one series of data per column or per row, no overlapping or intersection

Supported functions:

- SUM
- AVERAGE
- IF
- VLOOKUP
- HLOOKUP
- INDEX
- MATCH
- COUNT
- COUNTIF
- COUNTA
- UNIQUE
- COUNTIFS
- SUMIF
- SUMIFS
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
- AVERAGEIF
- AVERAGEIFS
- IFERROR
- SUBSTITUTE
- DATE
- DAY
- FILTER
- ROUNDDOWN
- ROUNDUP
- SORT
- SUMPRODUCT
- TEXTSPLIT
- TRIM
