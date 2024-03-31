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

- series_id: UUID
- worksheet: Worksheet
- series_header: str
- formulas: list[str]
- values: list[Union[int, str, float, bool]]
- header_location: HeaderLocation
- series_starting_cell: Cell
- series_length: int
- series_data_type: SeriesDataType

### SeriesRange

Attributes:

- series: list[Series]
- start_index: int
- end_index: int

### ASTGenerator

Attributes:

- formula_1_series: xlcalculator.ast_nodes.ASTNode
- formula_2_series: xlcalculator.ast_nodes.ASTNode

Methods:

- get_ast(index: int)

### ASTGeneratorPython

Attributes

Methods:

- get_ast_python(index: int)
