from formulas.tokens.operand import Range
from objects import Cell

from excel_utils import ExcelUtils
from typing import List


class SeriesImplementer:

    @staticmethod
    def get_cells_between(cell_start: Cell, cell_end: Cell) -> List[Cell]:
        cells = [
            Cell(row=row, column=column)
            for row in range(cell_start.row, cell_end.row + 1)
            for column in range(cell_start.column, cell_end.column + 1)
        ]
        return cells

    @staticmethod
    def create_cell(column: int, row: int, sheet_name: str) -> Cell:
        return Cell(
            column=column,
            row=row,
            sheet_name=sheet_name,
            value=None,
            value_type=None,
        )

    @staticmethod
    def get_cells_in_range(sheet_name: str, cell_range: str) -> List[Cell]:
        (
            cell_start_column,
            cell_start_row,
            cell_end_column,
            cell_end_row,
            _,
        ) = ExcelUtils.get_coordinates_from_range(cell_range)

        cell_start = SeriesImplementer.create_cell(
            cell_start_column, cell_start_row, sheet_name
        )
        cell_end = SeriesImplementer.create_cell(
            cell_end_column, cell_end_row, sheet_name
        )

        cells_in_range = SeriesImplementer.get_cells_between(cell_start, cell_end)
        return cells_in_range

    @staticmethod
    def implement_series(formula_ast_elements, series_mapping):
        """Implement series in formula AST"""
        for index, element in enumerate(formula_ast_elements):
            if isinstance(element, Range):
                cell_range_string = element.name
                cells_in_range = SeriesImplementer.get_cells_in_range(cell_range_string)
                series = series_mapping.get(cells_in_range[0])
                print(series)

        return formula_ast_elements
