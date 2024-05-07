import xlcalculator
import ast
from typing import List, Tuple
from objects import SeriesId


from dataclasses import dataclass


@dataclass
class SeriesRangeDelta:
    start_row_index_delta: int
    end_row_index_delta: int
    series_id_start_row_index_delta: int
    series_id_end_row_index_delta: int
    series_id_start_column_index_delta: int
    series_id_end_column_index_delta: int
    start_row_index: int
    end_row_index: int


class SeriesIdLoader:
    @staticmethod
    def load_series_id_from_string(series_id_string: str) -> SeriesId:
        sheet_name, series_header, series_header_cell_row, series_header_cell_column = (
            series_id_string.split("|")
        )
        return SeriesId(
            sheet_name=sheet_name,
            series_header=series_header,
            series_header_cell_row=int(series_header_cell_row),
            series_header_cell_column=int(series_header_cell_column),
        )


class DeltaCalculator:
    @staticmethod
    def load_series_ids(series_ids_strings: List[str]) -> List[SeriesId]:
        return [
            SeriesIdLoader.load_series_id_from_string(sid) for sid in series_ids_strings
        ]

    @staticmethod
    def calculate_index_deltas(
        indexes1: Tuple[int, int], indexes2: Tuple[int, int]
    ) -> Tuple[int, int]:
        return tuple(y - x for x, y in zip(indexes1, indexes2))


class FormulaGenerator:

    @staticmethod
    def traverse_and_replace(
        ast1: xlcalculator.ast_nodes.ASTNode, ast2: xlcalculator.ast_nodes.ASTNode
    ) -> xlcalculator.ast_nodes.ASTNode:
        if isinstance(ast1, xlcalculator.ast_nodes.RangeNode) and isinstance(
            ast2, xlcalculator.ast_nodes.RangeNode
        ):
            return FormulaGenerator.replace_range_node_with_formula(ast1, ast2)
        elif isinstance(ast1, xlcalculator.ast_nodes.FunctionNode) and isinstance(
            ast2, xlcalculator.ast_nodes.FunctionNode
        ):
            if ast1.token.tvalue == ast2.token.tvalue:
                modified_args = [
                    FormulaGenerator.traverse_and_replace(arg1, arg2)
                    for arg1, arg2 in zip(ast1.args, ast2.args)
                ]
                modified_function_node = xlcalculator.ast_nodes.FunctionNode(ast1.token)
                modified_function_node.args = modified_args
                return modified_function_node
        elif isinstance(ast1, xlcalculator.ast_nodes.OperatorNode) and isinstance(
            ast2, xlcalculator.ast_nodes.OperatorNode
        ):
            if ast1.token.tvalue == ast2.token.tvalue:
                modified_left = (
                    FormulaGenerator.traverse_and_replace(ast1.left, ast2.left)
                    if ast1.left and ast2.left
                    else None
                )
                modified_right = (
                    FormulaGenerator.traverse_and_replace(ast1.right, ast2.right)
                    if ast1.right and ast2.right
                    else None
                )
                modified_operator_node = xlcalculator.ast_nodes.OperatorNode(ast1.token)
                modified_operator_node.left = modified_left
                modified_operator_node.right = modified_right
                return modified_operator_node
        return ast1

    @staticmethod
    def replace_range_node_with_formula(
        node1: xlcalculator.ast_nodes.RangeNode, node2: xlcalculator.ast_nodes.RangeNode
    ) -> xlcalculator.ast_nodes.RangeNode:
        generic_formula = FormulaGenerator.get_generic_formula(node1, node2)
        print(type(generic_formula))
        print(generic_formula)
        return xlcalculator.ast_nodes.RangeNode(
            xlcalculator.tokenizer.f_token(
                tvalue=generic_formula, ttype="operand", tsubtype="range"
            )
        )

    @staticmethod
    def get_generic_formula(
        range_node1: xlcalculator.ast_nodes.RangeNode,
        range_node2: xlcalculator.ast_nodes.RangeNode,
    ) -> str:
        node1_value = range_node1.tvalue
        node2_value = range_node2.tvalue

        node1_tuple = ast.literal_eval(node1_value)
        node2_tuple = ast.literal_eval(node2_value)

        series_ids_1, row_indexes_1 = node1_tuple
        _, row_indexes_2 = node2_tuple

        row_index_deltas = DeltaCalculator.calculate_index_deltas(
            row_indexes_1, row_indexes_2
        )

        generic_formula = (series_ids_1, row_indexes_1, row_index_deltas)

        return str(generic_formula)
