# Classes

### ExcelFile

Attributes:

- file_path: str
- workbook: openpyxl.Workbook

### Cell

Attributes:

- column: int
- row: int
- coordinate: str

### CellRange

Attributes:

- start_cell: Cell
- end_cell: Cell

### Table

Attributes:

- name: str
- range: CellRange
- header_location: ['left', 'top']
- header_values: list[str]
