import xlcalculator
from objects import Cell, Worksheet, Series, SeriesRange
from coordinate_transformer import CoordinateTransformer


class SeriesMappingAccessor:
    def __init__(self, series_mapping: dict[Worksheet, dict[Cell, Series]]) -> None:
        self.series_mapping = series_mapping

    def get_series_from_cell(self, worksheet: Worksheet, cell: Cell):
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

    def get_cells_between(self, cell_start: Cell, cell_end: Cell) -> list[Cell]:
        cells = [
            Cell(row=row, column=column)
            for row in range(cell_start.row, cell_end.row + 1)
            for column in range(cell_start.column, cell_end.column + 1)
        ]
        return cells

    def get_series_range_from_cell_range(self, cell_range: str) -> SeriesRange:
        """cell_range is an Excel cell range as a string, eg. 'A1:B2' or 'A:B'"""

        (
            cell_start_column,
            cell_start_row,
            cell_end_column,
            cell_end_row,
            is_column_range,
        ) = CoordinateTransformer.get_coordinates_from_range(cell_range)

        cell_start = Cell(
            column=cell_start_column,
            row=cell_start_row,
            coordinate=None,
            value=None,
            value_type=None,
        )
        cell_end = Cell(
            column=cell_end_column,
            row=cell_end_row,
            coordinate=None,
            value=None,
            value_type=None,
        )

        cells_in_range = self.get_cells_between(cell_start, cell_end)

        worksheet = Worksheet(
            sheet_name=self.sheet_name, workbook_file_path=None, worksheet=None
        )

        series_list = [
            self.accessor.get_series_from_cell(worksheet, cell)
            for cell in cells_in_range
        ]

        filtered_series_list = [series for series in series_list if series is not None]

        series_range = SeriesRange(
            series=[item[1] for item in filtered_series_list],
            start_index=filtered_series_list[0][0],
            end_index=filtered_series_list[-1][0],
            is_column_range=is_column_range,
        )

        return series_range

    @staticmethod
    def get_series_ids_from_series_range(series_range: SeriesRange) -> str:
        series_ids = [
            str(series.series_id).replace("-", "") for series in series_range.series
        ]
        series_ids_unique = list(set(series_ids))

        series_ids_sorted = tuple(
            sorted(
                series_ids_unique, key=lambda x: (x.split("|")[-1], x.split("|")[-3])
            )
        )

        if series_range.is_column_range:
            return str(
                tuple(
                    [
                        series_ids_sorted,
                        tuple([None, None]),
                    ]
                )
            )

        return str(
            tuple(
                [
                    series_ids_sorted,
                    tuple([series_range.start_index, series_range.end_index]),
                ]
            )
        )

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

    def replace_range_nodes(self, ast):

        if isinstance(ast, xlcalculator.ast_nodes.RangeNode):

            if "!" in ast.tvalue:
                cell_range = self.extract_cell_ranges_from_string(ast.tvalue)
            else:
                cell_range = ast.tvalue

            series_range = self.get_series_range_from_cell_range(
                cell_range=cell_range,
            )
            series_ids = self.get_series_ids_from_series_range(series_range)

            return xlcalculator.ast_nodes.RangeNode(
                xlcalculator.tokenizer.f_token(
                    tvalue=series_ids, ttype="operand", tsubtype="range"
                )
            )
        elif isinstance(ast, xlcalculator.ast_nodes.FunctionNode):

            modified_args = [self.replace_range_nodes(arg) for arg in ast.args]
            modified_function_node = xlcalculator.ast_nodes.FunctionNode(ast.token)
            modified_function_node.args = modified_args
            return modified_function_node
        elif isinstance(ast, xlcalculator.ast_nodes.OperatorNode):
            modified_left = self.replace_range_nodes(ast.left) if ast.left else None
            modified_right = self.replace_range_nodes(ast.right) if ast.right else None
            modified_operator_node = xlcalculator.ast_nodes.OperatorNode(ast.token)
            modified_operator_node.left = modified_left
            modified_operator_node.right = modified_right
            return modified_operator_node
        else:
            return ast
