# Classes

### ExcelFile

Attributes:

- file_path: str
- workbook: openpyxl.Workbook

### Worksheet

Attributes:

- sheet_name: str
- workbook_name: Optional[str]
- worksheet: Optional[openpyxl.Worksheet]

### Cell

Attributes:

- column: int
- row: int
- coordinate: str

### CellRange

Attributes:

- start_cell: Cell
- end_cell: Cell

# HeaderLocation (ENUM)

Attributes:

- TOP: str
- LEFT: str

### Table

Attributes:

- name: str
- range: CellRange
- header_location: HeaderLocation
- header_values: list[str]

### SeriesDataType (ENUM)

- INT: str
- STR: str
- FLOAT: str
- BOOL: str

### Series

Attributes:

- series_header: str
- formulas: list[str]
- values: list[Union[int, str, float, bool]]
- header_location: HeaderLocation
- series_starting_cell: Cell
- series_length: int
- series_datat_type: SeriesDataType
