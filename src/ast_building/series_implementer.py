import xlcalculator
from objects import Cell, Worksheet, Series, SeriesRange


class SeriesImplementer:

    def __init__(self, series_mapping) -> None:
        self.series_mapping = series_mapping

    @staticmethod
    def get_series_from_cell_and_sheet_name(series_mapping, worksheet, cell):
        return series_mapping[worksheet][cell]

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
    def coordinate_from_string(cell_coordinate: str):
        """Convert Excel-style cell reference to numerical row and column indices."""
        column_str = "".join(filter(str.isalpha, cell_coordinate))
        row_str = "".join(filter(str.isdigit, cell_coordinate))

        # Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, ...)
        column = 0
        for char in column_str:
            column = column * 26 + (ord(char.upper()) - ord("A") + 1)

        # Convert row string to number
        row = int(row_str)

        return (column, row)

    @staticmethod
    def get_series_range_from_cell_range(
        series_mapping: dict, sheet_name: str, cell_range: str
    ) -> list[Series]:
        """cell_range is an Excel cell range as a string, eg. 'A1:B2'"""

        # Get the start and end cell coordinates from the cell range
        cell_start_coordinate, cell_end_coordinate = cell_range.split(":")

        # Convert cell coordinates to row and column
        cell_start_column, cell_start_row = SeriesImplementer.coordinate_from_string(
            cell_start_coordinate
        )
        cell_end_column, cell_end_row = SeriesImplementer.coordinate_from_string(
            cell_end_coordinate
        )

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

        series_range = SeriesRange(
            series=[item[1] for item in series_list],
            start_index=series_list[0][0],
            end_index=series_list[-1][0],
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
    def get_series_uuids_from_series_range(series_range: SeriesRange):
        series_ids = [
            str(series.series_id).replace("-", "") for series in series_range.series
        ]
        series_ids_unique = list(set(series_ids))
        return f'{"_".join(series_ids_unique)}_{series_range.start_index}_{series_range.end_index}'

    def replace_range_nodes(self, ast):
        if isinstance(ast, xlcalculator.ast_nodes.RangeNode):
            # Accessing series_range from the instance method directly
            series_range = self.get_series_range_from_cell_range(
                series_mapping=self.series_mapping,  # Accessed via self
                sheet_name="Sheet1",  # Assuming 'Sheet1' is intended or dynamically determined elsewhere
                cell_range=ast.tvalue,
            )
            series_uuids = self.get_series_uuids_from_series_range(
                series_range  # Accessed via self
            )
            # Assuming xlcalculator.tokenizer.f_token exists and works as shown
            return xlcalculator.ast_nodes.RangeNode(
                xlcalculator.tokenizer.f_token(
                    tvalue=series_uuids, ttype="operand", tsubtype="range"
                )
            )
        elif isinstance(ast, xlcalculator.ast_nodes.FunctionNode):
            modified_args = [
                self.replace_range_nodes(arg)
                for arg in ast.args  # Corrected to use self for instance method call
            ]
            modified_function_node = xlcalculator.ast_nodes.FunctionNode(ast.token)
            modified_function_node.args = modified_args
            return modified_function_node
        elif isinstance(ast, xlcalculator.ast_nodes.OperatorNode):
            modified_left = (
                self.replace_range_nodes(ast.left)
                if ast.left
                else None  # Corrected to use self for instance method call
            )
            modified_right = (
                self.replace_range_nodes(ast.right)
                if ast.right
                else None  # Corrected to use self for instance method call
            )
            modified_operator_node = xlcalculator.ast_nodes.OperatorNode(ast.token)
            modified_operator_node.left = modified_left
            modified_operator_node.right = modified_right
            return modified_operator_node
        else:
            return ast
