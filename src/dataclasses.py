from dataclasses import dataclass
from typing import Optional
import openpyxl


@dataclass
class ExcelFile:
    file_path: str
    workbook: openpyxl.Workbook


@dataclass
class Worksheet:
    sheet_name: str
    workbook_name: Optional[str]
    worksheet: Optional[openpyxl.Worksheet]


@dataclass
class Cell:
    column: int
    row: int
    coordinate: str


@dataclass
class CellRange:
    start_cell: "Cell"
    end_cell: "Cell"
