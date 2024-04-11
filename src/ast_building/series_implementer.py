import xlcalculator
from objects import Cell, Worksheet, Series, SeriesRange
from typing import Tuple, Optional


class SeriesImplementer:

    def __init__(
        self, series_mapping: dict[Worksheet, dict[Cell, Series]], sheet_name: str
    ) -> None:
        self.series_mapping = series_mapping
        self.sheet_name = sheet_name

    @staticmethod
    def get_series_from_cell_and_sheet_name(series_mapping, worksheet, cell):
        try:
            return series_mapping[worksheet][cell]
        except KeyError:
            return None

    @staticmethod
    def get_cells_between(cell_start: Cell, cell_end: Cell):
        """cell_start and cell_end as inputs. Get a list of all cells between these two cells."""
        cells = []
        for row in range(cell_start.row, cell_end.row + 1):
            for column in range(cell_start.column, cell_end.column + 1):
                cells.append(
                    Cell(
                        row=row,
                        column=column,
                        coordinate=None,
                        value=None,
                        value_type=None,
                    )
                )
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

        cell_start_column, cell_start_row = SeriesImplementer.get_coordinates_from_cell(
            cell_start
        )
        cell_end_column, cell_end_row = SeriesImplementer.get_coordinates_from_cell(
            cell_end
        )

        return (
            cell_start_column,
            cell_start_row,
            cell_end_column,
            cell_end_row,
            is_column_range,
        )

    @staticmethod
    def get_series_range_from_cell_range(
        series_mapping: dict, sheet_name: str, cell_range: str
    ) -> list[Series]:
        """cell_range is an Excel cell range as a string, eg. 'A1:B2' or 'A:B'"""

        (
            cell_start_column,
            cell_start_row,
            cell_end_column,
            cell_end_row,
            is_column_range,
        ) = SeriesImplementer.get_coordinates_from_range(cell_range)

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

        cells_in_range = SeriesImplementer.get_cells_between(cell_start, cell_end)

        worksheet = Worksheet(
            sheet_name=sheet_name, workbook_file_path=None, worksheet=None
        )

        series_list = [
            SeriesImplementer.get_series_from_cell_and_sheet_name(
                series_mapping=series_mapping, worksheet=worksheet, cell=cell
            )
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
    def serialise_ast_to_formula(ast):
        if isinstance(ast, xlcalculator.ast_nodes.RangeNode):
            value = ast.tvalue.strip("[]")
            return f"{value}"
        elif isinstance(ast, xlcalculator.ast_nodes.FunctionNode):
            args = ", ".join(
                SeriesImplementer.serialise_ast_to_formula(arg) for arg in ast.args
            )
            return f"{ast.tvalue}({args})"
        elif isinstance(ast, xlcalculator.ast_nodes.OperatorNode):
            left = (
                SeriesImplementer.serialise_ast_to_formula(ast.left) if ast.left else ""
            )
            right = (
                SeriesImplementer.serialise_ast_to_formula(ast.right)
                if ast.right
                else ""
            )
            return f"({left} {ast.tvalue} {right})".strip()
        elif (
            isinstance(ast, xlcalculator.ast_nodes.OperandNode)
            and ast.tsubtype == "text"
        ):
            return f'"{ast.tvalue}"'
        else:
            return str(ast.tvalue)

    @staticmethod
    def get_series_ids_from_series_range(series_range: SeriesRange) -> str:
        series_ids = [
            str(series.series_id).replace("-", "") for series in series_range.series
        ]
        series_ids_unique = list(set(series_ids))

        # sort series ids by column (second last index), then row (last index): Eg. ('Sheet1|col_2|2|2', 'Sheet1|col_1|2|1', 'Sheet1|col_3|4|2') ==> ('Sheet1|col_2|2|1', 'Sheet1|col_1|2|2', 'Sheet1|col_3|4|2')
        series_ids_sorted = tuple(
            sorted(
                series_ids_unique, key=lambda x: (x.split("|")[-1], x.split("|")[-3])
            )
        )

        if series_range.is_column_range:
            return str(tuple(series_ids_sorted))

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
                series_mapping=self.series_mapping,
                sheet_name=self.sheet_name,
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
