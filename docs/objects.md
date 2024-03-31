# Objects

### excel_raw_file_path

- Type: str
- Description: Path to raw Excel

### excel_raw

- Type: ExcelFile
- Description: Loaded raw Excel

### excel_reduced_file_path

- Type: str
- Description: Path to Excel with only two rows per table

### excel_reduced

- Type: ExcelFile
- Description: Loaded raw Excel file

### excel_reduced_clean

- Type: ExcelFile
- Description: Excel with only two rows per table with formulas cleaned

### extracted_tables

- Type: Dict[Worksheet, list[Table]]
- Description: Located tables inside an Excel with table metadata
- Example:

```python
{
  Worksheet(sheet_name="Sheet1"): [
    Table(name="Sheet1_1",
          range=CellRange(start_cell=Cell(column=2, row=12, coordinate='B12'),
                          end_cell=Cell(column=5, row=20, coordinate='E20')),
          header_location="top",
          header_values=["Column 1", "Column 2", "Column 3", "Column 4"]),
    Table(name="Sheet1_2",
          range=CellRange(start_cell=Cell(column=7, row=5, coordinate='G5'),
                          end_cell=Cell(column=10, row=15, coordinate='J15')),
          header_location="left",
          header_values=["Row 1", "Row 2", "Row 3", "Row 4", "Row 5",
                         "Row 6", "Row 7", "Row 8", "Row 9", "Row 10", "Row 11"])
  ],
  Worksheet(sheet_name="Sheet2"): [
    Table(name="Sheet2_1",
          range=CellRange(start_cell=Cell(column=1, row=1, coordinate='A1'),
                          end_cell=Cell(column=3, row=10, coordinate='C10')),
          header_location="top",
          header_values=["Data 1", "Data 2", "Data 3"]),
    Table(name="Sheet2_2",
          range=CellRange(start_cell=Cell(column=4, row=3, coordinate='D3'),
                          end_cell=Cell(column=8, row=12, coordinate='H12')),
          header_location="top",
          header_values=["Info A", "Info B", "Info C", "Info D", "Info E"])
  ]
}
```

### series_data

- Type: dict[str, dict[CellRange, Series]]
- Description: Collection of series
- Example:

```python
{
    Worksheet(sheet_name="Sheet1"): {
        CellRange(coordinates='C12:D12'): Series(
            series_header='horizontal_column_1',
            formulas=['=B3', '=C3'],
            values=[1, 2],
            header_location=HeaderLocation.LEFT,
            series_starting_cell=Cell(row=12, column=3),
            series_length=2,
            data_type=SeriesDataType.INT
        ),
        CellRange(coordinates='C13:D13'): Series(
            series_header='horizontal_column_2',
            formulas=['=B4', '=C4'],
            values=[3, 4],
            header_location=HeaderLocation.LEFT,
            series_starting_cell=Cell(row=13, column=3),
            series_length=2,
            data_type=SeriesDataType.INT
        ),
        CellRange(coordinates='B3:B4'): Series(
            series_header='col_1',
            formulas=[None, None],
            values=[1, 3],
            header_location=HeaderLocation.TOP,
            series_starting_cell=Cell(row=3, column=2),
            series_length=2,
            data_type=SeriesDataType.INT
        )
    },
    Worksheet(sheet_name="Sheet2"): {
        CellRange(coordinates='A11:A12'): Series(
            series_header='col1',
            formulas=['=B2', '=B3'],
            values=[1, 4],
            header_location=HeaderLocation.TOP,
            series_starting_cell=Cell(row=11, column=1),
            series_length=2,
            data_type=SeriesDataType.INT
        ),
        CellRange(coordinates='B11:B12'): Series(
            series_header='col2',
            formulas=['=C2', '=C3'],
            values=[2, 5],
            header_location=HeaderLocation.TOP,
            series_starting_cell=Cell(row=11, column=2),
            series_length=2,
            data_type=SeriesDataType.INT
        ),
        CellRange(coordinates='B2:C2'): Series(
            series_header='horizontal_col_1',
            formulas=[None, None],
            values=[1, 2],
            header_location=HeaderLocation.LEFT,
            series_starting_cell=Cell(row=2, column=2),
            series_length=2,
            data_type=SeriesDataType.INT
        ),
        CellRange(coordinates='B3:C3'): Series(
            series_header='horizontal_col_2',
            formulas=[None, None],
            values=[4, 5],
            header_location=HeaderLocation.LEFT,
            series_starting_cell=Cell(row=3, column=2),
            series_length=2,
            data_type=SeriesDataType.INT
        )
    }
}

```

### formula_1

- Type: str
- Description: First formula inside a series
- Example: "=SUM(A1:B3)"

### formula_2

- Type: str
- Description: Second formula inside a series
- Example: "=SUM(A1:B3)"

### formula_1_ast

- Type: xlcalculator.ast_nodes.ASTNode
- Description: Represents the starting node of the AST for formula_1, which can be traversed
- Example:

```python
xlcalculator.ast_nodes.FunctionNode(tvalue='SUM',ttype='function')
```

### formula_2_ast

- Type: xlcalculator.ast_nodes.ASTNode
- Description: Represents the starting node of the AST for formula_2, which can be traversed
- Example:

```python
xlcalculator.ast_nodes.FunctionNode(tvalue='SUM',ttype='function')
```
