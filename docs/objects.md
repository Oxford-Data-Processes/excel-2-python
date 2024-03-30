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

- Type: Dict[str, list[Table]]
- Example:

```python
{
  "Sheet1": [
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
  "Sheet2": [
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

- Description: Located tables inside an Excel with table metadata

### series_data

- Type:
