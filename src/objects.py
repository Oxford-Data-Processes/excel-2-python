from dataclasses import dataclass
from typing import Optional, List, Union
from uuid import UUID
from enum import Enum

import openpyxl


@dataclass
class ExcelFile:
    file_path: str
    workbook: openpyxl.Workbook


@dataclass
class Worksheet:
    sheet_name: str
    workbook_name: Optional[str]
    worksheet: Optional[openpyxl.worksheet.worksheet.Worksheet]


@dataclass
class Cell:
    column: int
    row: int
    coordinate: str


@dataclass
class CellRange:
    start_cell: "Cell"
    end_cell: "Cell"


class HeaderLocation(Enum):
    TOP = "top"
    LEFT = "left"


@dataclass
class Table:
    name: str
    range: CellRange
    header_location: HeaderLocation
    header_values: List[str]


class SeriesDataType(Enum):
    INT = "int"
    STR = "str"
    FLOAT = "float"
    BOOL = "bool"


@dataclass
class Series:
    series_id: UUID
    worksheet: Worksheet
    series_header: str
    formulas: List[str]
    values: List[Union[int, str, float, bool]]
    header_location: HeaderLocation
    series_starting_cell: Cell
    series_length: int
    series_data_type: SeriesDataType


@dataclass
class SeriesRange:
    series: List[Series]
    start_index: int
    end_index: int