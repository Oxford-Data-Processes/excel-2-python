from dataclasses import dataclass, field
from typing import Optional, List, Union, Dict
from enum import Enum
from excel_utils import ExcelUtils

import openpyxl


@dataclass
class ExcelFile:
    file_path: str
    workbook_with_formulas: openpyxl.Workbook
    workbook_with_values: openpyxl.Workbook


@dataclass(frozen=True)
class Worksheet:
    sheet_name: str
    workbook_file_path: Optional[str] = None
    worksheet: Optional[openpyxl.worksheet.worksheet.Worksheet] = None


@dataclass(frozen=True)
class Cell:
    column: int
    row: int
    coordinate: Optional[str] = None
    sheet_name: Optional[str] = None
    value: Optional[Union[int, str, float, bool]] = None
    value_type: Optional[str] = None
    formula: Optional[str] = None

    def __post_init__(self):
        object.__setattr__(self, "coordinate", self.calculate_coordinate())

    def calculate_coordinate(self) -> str:
        if self.column and self.row:
            return f"{ExcelUtils.get_column_letter_from_number(self.column)}{self.row}"
        return None


@dataclass
class WorkbookData:
    data: Dict[str, Dict[str, Cell]] = field(default_factory=dict)

    def add_sheet_data(self, sheet_name: str, sheet_data: Dict[str, Cell]):
        """Add or update the data for a specific worksheet."""
        self.data[sheet_name] = sheet_data

    def get_sheet_data(self, sheet_name: str) -> Optional[Dict[str, Cell]]:
        """Retrieve the data for a specific worksheet."""
        return self.data.get(sheet_name)


@dataclass
class CellRange:
    start_cell: "Cell"
    end_cell: "Cell"

    def __str__(self):
        if self.start_cell == self.end_cell:
            return f"{self.start_cell.sheet_name}!{self.start_cell.coordinate}"
        return f"{self.start_cell.sheet_name}!{self.start_cell.coordinate}:{self.end_cell.coordinate}"


@dataclass
class Column:
    column_number: int
    column_letter: Optional[str] = None


@dataclass
class CellRangeColumn:
    start_column: Column
    end_column: Column
    sheet_name: str

    def __str__(self):
        return f"{self.sheet_name}!{self.start_column.column_letter}:{self.end_column.column_letter}"


@dataclass
class Table:
    name: str
    range: CellRange
    header_values: Optional[List[Union[int, str, float, bool]]] = None


@dataclass
class LocatedTables:
    worksheet: Worksheet
    tables: List[Table] = field(default_factory=list)


class SeriesDataType(Enum):
    INT = "int"
    STR = "str"
    FLOAT = "float"
    BOOL = "bool"
    TIME = "time"


@dataclass(frozen=True)
class SeriesId:
    sheet_name: str
    series_header: str
    series_header_cell_row: int
    series_header_cell_column: int

    def __str__(self):
        return "|".join(
            [
                self.sheet_name,
                self.series_header,
                str(self.series_header_cell_row),
                str(self.series_header_cell_column),
            ]
        )


@dataclass
class Series:
    series_id: SeriesId
    worksheet: Worksheet
    series_header: str
    formulas: List[str]
    values: List[Union[int, str, float, bool]]
    series_starting_cell: Cell
    series_length: int
    series_data_type: SeriesDataType


@dataclass
class SeriesRange:
    series: List[Series]
    start_index: Optional[int]
    end_index: Optional[int]
    is_column_range: bool = False
