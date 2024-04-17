import xlcalculator
import ast

from objects import Cell, HeaderLocation, CellRange, Column, CellRangeColumn

from excel_utils import ExcelUtils
from ast_transformation.formula_generator import SeriesIdLoader


class CellRangeImplementer:

    def __init__(self, series_dict):
        self.series_dict = series_dict

    def merge_cell_ranges(self, cell_ranges):

        min_row = min(cell_range.start_cell.row for cell_range in cell_ranges)
        min_column = min(cell_range.start_cell.column for cell_range in cell_ranges)
        max_row = max(cell_range.end_cell.row for cell_range in cell_ranges)
        max_column = max(cell_range.end_cell.column for cell_range in cell_ranges)

        start_cell = Cell(
            column=min_column,
            row=min_row,
            coordinate=ExcelUtils.get_coordinate_from_column_and_row(
                min_column, min_row
            ),
            sheet_name=cell_ranges[0].start_cell.sheet_name,
        )
        end_cell = Cell(
            column=max_column,
            row=max_row,
            coordinate=ExcelUtils.get_coordinate_from_column_and_row(
                max_column, max_row
            ),
            sheet_name=cell_ranges[0].start_cell.sheet_name,
        )

        return CellRange(start_cell=start_cell, end_cell=end_cell)

    def create_cell_range_top_header(
        self, start_index, end_index, cell_row, cell_column, sheet_name
    ):

        return CellRange(
            start_cell=Cell(
                row=cell_row + start_index,
                column=cell_column,
                coordinate=ExcelUtils.get_coordinate_from_column_and_row(
                    cell_column, cell_row + start_index
                ),
                sheet_name=sheet_name,
            ),
            end_cell=Cell(
                row=cell_row + start_index,
                column=cell_column,
                coordinate=ExcelUtils.get_coordinate_from_column_and_row(
                    cell_column, cell_row + start_index
                ),
                sheet_name=sheet_name,
            ),
        )

    def create_cell_range_left_header(
        self, start_index, end_index, cell_row, cell_column, sheet_name
    ):
        return CellRange(
            start_cell=Cell(
                row=cell_row,
                column=cell_column + start_index,
                coordinate=ExcelUtils.get_coordinate_from_column_and_row(
                    cell_column + start_index, cell_row
                ),
                sheet_name=sheet_name,
            ),
            end_cell=Cell(
                row=cell_row,
                column=cell_column + start_index,
                coordinate=ExcelUtils.get_coordinate_from_column_and_row(
                    cell_column + start_index, cell_row
                ),
                sheet_name=sheet_name,
            ),
        )

    def get_cell_range_from_series_tuple(self, series_tuple):

        series_ids_string, indexes = series_tuple
        series_start_index, series_end_index = indexes

        if series_start_index is None and series_end_index is None:
            return self.process_series_columns(series_ids_string)

        return self.process_series_cells(
            series_ids_string, series_start_index, series_end_index
        )

    def process_series_columns(self, series_ids_string):
        column_values = []

        sheet_name = SeriesIdLoader.load_series_id_from_string(
            series_ids_string[0]
        ).sheet_name
        for series_id_string in series_ids_string:
            column_value = self.get_column_from_series_id(series_id_string)
            column_values.append(column_value)

        sorted_column_values = sorted(column_values, key=lambda x: x.column_number)
        return CellRangeColumn(
            start_column=sorted_column_values[0],
            end_column=sorted_column_values[-1],
            sheet_name=sheet_name,
        )

    def get_column_from_series_id(self, series_id_string):
        series_id = SeriesIdLoader.load_series_id_from_string(series_id_string)
        sheet_name = series_id.sheet_name
        series_list = self.series_dict.get(sheet_name)

        for series in series_list:
            if series.series_id == series_id:
                column_value = series.series_starting_cell.column
                return Column(
                    column_number=column_value,
                    column_letter=ExcelUtils.get_column_letter_from_number(
                        column_value
                    ),
                )

    def process_series_cells(
        self, series_ids_string, series_start_index, series_end_index
    ):

        cell_ranges = []

        for series_id_string in series_ids_string:

            cell_range = self.get_cell_range_for_series_id(
                series_id_string, series_start_index, series_end_index
            )
            cell_ranges.append(cell_range)

        return self.merge_cell_ranges(cell_ranges)

    def get_cell_range_for_series_id(
        self, series_id_string, series_start_index, series_end_index
    ):
        series_id = SeriesIdLoader.load_series_id_from_string(series_id_string)
        sheet_name = series_id.sheet_name
        series_list = self.series_dict.get(sheet_name)

        for series in series_list:
            if series.series_id == series_id:
                return self.create_cell_range(
                    series, series_start_index, series_end_index, sheet_name
                )

    def create_cell_range(
        self, series, series_start_index, series_end_index, sheet_name
    ):
        cell_value = series.series_starting_cell
        cell_row = cell_value.row
        cell_column = cell_value.column

        if series.header_location == HeaderLocation.TOP:
            return self.create_cell_range_top_header(
                series_start_index, series_end_index, cell_row, cell_column, sheet_name
            )
        elif series.header_location == HeaderLocation.LEFT:
            return self.create_cell_range_left_header(
                series_start_index, series_end_index, cell_row, cell_column, sheet_name
            )
        else:
            raise Exception("Header location is not valid")

    def update_ast(self, ast):
        if isinstance(ast, xlcalculator.ast_nodes.RangeNode):
            return self.replace_range_node(ast)
        elif isinstance(ast, xlcalculator.ast_nodes.FunctionNode):
            return self.replace_function_node(ast)
        elif isinstance(ast, xlcalculator.ast_nodes.OperatorNode):
            return self.replace_operator_node(ast)
        return ast

    def replace_range_node(self, node):

        series_tuple = ast.literal_eval(node.tvalue)
        cell_range = self.get_cell_range_from_series_tuple(series_tuple)

        return xlcalculator.ast_nodes.RangeNode(
            xlcalculator.tokenizer.f_token(
                tvalue=cell_range, ttype="operand", tsubtype="range"
            )
        )

    def replace_function_node(self, node):
        modified_args = [self.update_ast(arg) for arg in node.args]
        modified_function_node = xlcalculator.ast_nodes.FunctionNode(node.token)
        modified_function_node.args = modified_args
        return modified_function_node

    def replace_operator_node(self, node):
        modified_left = self.update_ast(node.left) if node.left else None
        modified_right = self.update_ast(node.right) if node.right else None
        modified_operator_node = xlcalculator.ast_nodes.OperatorNode(node.token)
        modified_operator_node.left = modified_left
        modified_operator_node.right = modified_right
        return modified_operator_node
