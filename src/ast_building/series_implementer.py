import xlcalculator
from objects import Cell, Worksheet, Series, SeriesRange
from excel_utils import ExcelUtils

from typing import Optional


class SeriesMappingAccessor:
    def __init__(self, series_mapping: dict[Worksheet, dict[Cell, Series]]) -> None:
        self.series_mapping = series_mapping

    def get_series_from_cell(
        self, worksheet: Worksheet, cell: Cell
    ) -> Optional[Series]:
        try:
            return self.series_mapping[worksheet][cell]
        except KeyError:
            return None


class SeriesImplementer:

    def __init__(
        self, series_mapping: dict[Worksheet, dict[Cell, Series]], sheet_name: str
    ) -> None:
        self.sheet_name = sheet_name
        self.accessor = SeriesMappingAccessor(series_mapping)

    @staticmethod
    def get_cells_between(cell_start: Cell, cell_end: Cell) -> list[Cell]:
        cells = [
            Cell(row=row, column=column)
            for row in range(cell_start.row, cell_end.row + 1)
            for column in range(cell_start.column, cell_end.column + 1)
        ]
        return cells

    def create_cell(self, column: int, row: int, sheet_name: str) -> Cell:
        return Cell(
            column=column,
            row=row,
            sheet_name=sheet_name,
            coordinate=None,
            value=None,
            value_type=None,
        )

    def get_series_list(
        self, cells_in_range: list[Cell], worksheet: Worksheet
    ) -> list[Series]:
        return [
            self.accessor.get_series_from_cell(worksheet, cell)
            for cell in cells_in_range
        ]

    def filter_series_list(self, series_list: list[Series]) -> list[Series]:
        return [series for series in series_list if series is not None]

    def create_series_range(
        self, filtered_series_list: list[Series], is_column_range: bool
    ) -> SeriesRange:
        return SeriesRange(
            series=[item[1] for item in filtered_series_list],
            start_index=filtered_series_list[0][0],
            end_index=filtered_series_list[-1][0],
            is_column_range=is_column_range,
        )

    def get_series_range_from_cell_range(
        self, sheet_name: str, cell_range: str
    ) -> SeriesRange:
        """cell_range is an Excel cell range as a string, eg. 'A1:B2' or 'A:B'"""

        (
            cell_start_column,
            cell_start_row,
            cell_end_column,
            cell_end_row,
            is_column_range,
        ) = ExcelUtils.get_coordinates_from_range(cell_range)

        cell_start = self.create_cell(cell_start_column, cell_start_row, sheet_name)
        cell_end = self.create_cell(cell_end_column, cell_end_row, sheet_name)

        cells_in_range = self.get_cells_between(cell_start, cell_end)
        worksheet = Worksheet(sheet_name=sheet_name)
        series_list = self.get_series_list(cells_in_range, worksheet)
        filtered_series_list = self.filter_series_list(series_list)
        series_range = self.create_series_range(filtered_series_list, is_column_range)

        return series_range

    @staticmethod
    def get_series_ids_from_series_range(series_range: SeriesRange) -> str:
        series_ids = [series.series_id for series in series_range.series]
        series_ids_unique = list(set(series_ids))

        series_ids_sorted = sorted(
            series_ids_unique,
            key=lambda x: (x.series_header_cell_column, x.series_header),
        )

        series_ids_sorted_strings = tuple([str(sid) for sid in series_ids_sorted])

        if series_range.is_column_range:
            return str(
                tuple(
                    [
                        series_ids_sorted_strings,
                        tuple([None, None]),
                    ]
                )
            )

        return str(
            tuple(
                [
                    series_ids_sorted_strings,
                    tuple([series_range.start_index, series_range.end_index]),
                ]
            )
        )

    def update_ast(
        self, ast: xlcalculator.ast_nodes.ASTNode
    ) -> xlcalculator.ast_nodes.ASTNode:
        if isinstance(ast, xlcalculator.ast_nodes.RangeNode):
            return self.replace_range_node(ast)
        elif isinstance(ast, xlcalculator.ast_nodes.FunctionNode):
            return self.replace_function_node(ast)
        elif isinstance(ast, xlcalculator.ast_nodes.OperatorNode):
            return self.replace_operator_node(ast)
        return ast

    def replace_range_node(
        self, node: xlcalculator.ast_nodes.RangeNode
    ) -> xlcalculator.ast_nodes.RangeNode:
        if "!" in node.tvalue:
            sheet_name, cell_range = ExcelUtils.extract_cell_ranges_from_string(
                node.tvalue
            )
        else:
            cell_range = node.tvalue
            sheet_name = self.sheet_name

        series_range = self.get_series_range_from_cell_range(
            cell_range=cell_range, sheet_name=sheet_name
        )

        series_ids = self.get_series_ids_from_series_range(series_range)

        return xlcalculator.ast_nodes.RangeNode(
            xlcalculator.tokenizer.f_token(
                tvalue=series_ids, ttype="operand", tsubtype="range"
            )
        )

    def replace_function_node(
        self, node: xlcalculator.ast_nodes.FunctionNode
    ) -> xlcalculator.ast_nodes.FunctionNode:
        modified_args = [self.update_ast(arg) for arg in node.args]
        modified_function_node = xlcalculator.ast_nodes.FunctionNode(node.token)
        modified_function_node.args = modified_args
        return modified_function_node

    def replace_operator_node(
        self, node: xlcalculator.ast_nodes.OperatorNode
    ) -> xlcalculator.ast_nodes.OperatorNode:
        modified_left = self.update_ast(node.left) if node.left else None
        modified_right = self.update_ast(node.right) if node.right else None
        modified_operator_node = xlcalculator.ast_nodes.OperatorNode(node.token)
        modified_operator_node.left = modified_left
        modified_operator_node.right = modified_right
        return modified_operator_node
