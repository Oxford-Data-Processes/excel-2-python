from typing import Tuple, Optional
from objects import Cell


class ExcelUtils:

    @staticmethod
    def get_cells_between(cell_start: Cell, cell_end: Cell) -> list[Cell]:
        cells = [
            Cell(row=row, column=column)
            for row in range(cell_start.row, cell_end.row + 1)
            for column in range(cell_start.column, cell_end.column + 1)
        ]
        return cells

    @staticmethod
    def get_coordinates_from_cell(cell_coordinate: str):

        column_str = "".join(filter(str.isalpha, cell_coordinate))
        row_str = "".join(filter(str.isdigit, cell_coordinate))

        column = 0
        for char in column_str:
            column = column * 26 + (ord(char.upper()) - ord("A") + 1)

        row = int(row_str)

        return (column, row)

    @staticmethod
    def extract_cell_ranges_from_string(cell_range_string: str):
        if ":" in cell_range_string:
            cell_range_start, cell_range_end = cell_range_string.split(":")
        else:
            cell_range_start = cell_range_string
            cell_range_end = cell_range_string

        if "!" in cell_range_start:
            _, cell_range_start = cell_range_start.split("!")
        if "!" in cell_range_end:
            _, cell_range_end = cell_range_end.split("!")

        return f"{cell_range_start}:{cell_range_end}"

    @staticmethod
    def get_coordinates_from_range(
        cell_range: str,
    ) -> Tuple[int, Optional[int], int, Optional[int], bool]:
        """Convert Excel-style cell range reference Eg. "A1:B3" or "A1" to numerical row and column indices."""

        if not ":" in cell_range:
            cell_range = cell_range + ":" + cell_range

        cell_start, cell_end = cell_range.split(":")

        is_column_range = False

        def check_is_column(cell_str: str):
            return cell_str.isalpha()

        if check_is_column(cell_start) and check_is_column(cell_end):
            cell_start = cell_start + "1"
            cell_end = cell_end + "3"
            is_column_range = True

        cell_start_column, cell_start_row = ExcelUtils.get_coordinates_from_cell(
            cell_start
        )
        cell_end_column, cell_end_row = ExcelUtils.get_coordinates_from_cell(cell_end)

        return (
            cell_start_column,
            cell_start_row,
            cell_end_column,
            cell_end_row,
            is_column_range,
        )