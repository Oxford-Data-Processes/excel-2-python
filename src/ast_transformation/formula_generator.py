import xlcalculator
import ast
from typing import List
from objects import Series


class SeriesRangeDelta:
    def __init__(
        self,
        start_row_index_delta,
        end_row_index_delta,
        start_column_index_delta,
        end_column_index_delta,
        start_row_index,
        end_row_index,
    ):
        self.start_row_index_delta = start_row_index_delta
        self.end_row_index_delta = end_row_index_delta
        self.start_column_index_delta = start_column_index_delta
        self.end_column_index__delta = end_column_index_delta
        self.start_row_index = start_row_index
        self.end_row_index = end_row_index


class ASTGenerator:
    def __init__(
        self,
        formula_1_ast_series: xlcalculator.ast_nodes.ASTNode,
        formula_2_ast_series: xlcalculator.ast_nodes.ASTNode,
        series_list: List[Series],
    ):
        self.formula_1_ast_series = formula_1_ast_series
        self.formula_2_ast_series = formula_2_ast_series
        self.series_list = series_list

    @staticmethod
    def extract_tuples(node_value: str):
        return ast.literal_eval(node_value)

    def get_delta_between_nodes(self, node1_value: str, node2_value: str):
        node1_tuple = self.extract_tuples(node1_value)
        node2_tuple = self.extract_tuples(node2_value)

        if node1_tuple[1] == (None, None) or node2_tuple[1] == (None, None):
            return None
        else:
            return self.calculate_deltas(node1_tuple, node2_tuple)

    @staticmethod
    def calculate_deltas(
        node1_tuple: tuple[tuple[str], tuple[int, int]],
        node2_tuple: tuple[tuple[str], tuple[int, int]],
    ) -> SeriesRangeDelta:
        node1_series_ids, node1_row_indexes = node1_tuple
        node2_series_ids, node2_row_indexes = node2_tuple

        start_row_index_delta = node2_row_indexes[0] - node1_row_indexes[0]
        end_row_index_delta = node2_row_indexes[1] - node1_row_indexes[1]
        start_column_index_delta = int(node2_series_ids[0].split("|")[-1]) - int(
            node1_series_ids[0].split("|")[-1]
        )
        end_column_index_delta = int(node2_series_ids[-1].split("|")[-1]) - int(
            node1_series_ids[-1].split("|")[-1]
        )

        return SeriesRangeDelta(
            start_row_index_delta,
            end_row_index_delta,
            start_column_index_delta,
            end_column_index_delta,
            node1_row_indexes[0],
            node1_row_indexes[-1],
        )

    def apply_delta_to_range_node(
        self,
        node1: xlcalculator.ast_nodes.ASTNode,
        node2: xlcalculator.ast_nodes.ASTNode,
        n: int,
    ):
        if isinstance(node1, xlcalculator.ast_nodes.RangeNode) and isinstance(
            node2, xlcalculator.ast_nodes.RangeNode
        ):
            return self.process_range_node(node1, node2, n)
        elif hasattr(node1, "args") and isinstance(
            node1, xlcalculator.ast_nodes.FunctionNode
        ):
            return self.process_function_node(node1, node2, n)
        elif hasattr(node1, "left") and isinstance(
            node1, xlcalculator.ast_nodes.OperatorNode
        ):
            return self.process_operator_node(node1, node2, n)
        else:
            return node1

    def process_range_node(self, node1, node2, n):
        series_range_delta = self.get_delta_between_nodes(node1.tvalue, node2.tvalue)
        if series_range_delta:

            start_row_index_delta = series_range_delta.start_row_index_delta
            end_row_index_delta = series_range_delta.end_row_index_delta
            start_column_index_delta = series_range_delta.start_column_index_delta
            end_column_index_delta = series_range_delta.end_column_index__delta
            start_row_index = series_range_delta.start_row_index
            end_row_index = series_range_delta.end_row_index
            return self.update_range_node(
                node1,
                start_row_index_delta,
                end_row_index_delta,
                start_column_index_delta,
                end_column_index_delta,
                start_row_index,
                end_row_index,
                n,
            )
        return node1

    def update_range_node(
        self,
        node1,
        start_row_index_delta,
        end_row_index_delta,
        start_column_index_delta,
        end_column_index_delta,
        start_row_index,
        end_row_index,
        n,
    ):

        series_ids = self.extract_tuples(node1.tvalue)[0]
        new_series_ids = [
            self.add_column_delta_to_series_id(sid, start_column_index_delta * (n - 1))
            for sid in series_ids
        ]

        new_tvalue = str(
            (
                tuple(new_series_ids),
                (
                    start_row_index + start_row_index_delta * (n - 1),
                    end_row_index + end_row_index_delta * (n - 1),
                ),
            )
        )
        return xlcalculator.ast_nodes.RangeNode(
            xlcalculator.tokenizer.f_token(
                tvalue=new_tvalue, ttype="operand", tsubtype="range"
            )
        )

    def add_column_delta_to_series_id(self, series_id, column_delta, series_list):
        sheet_name, series_header, start_column_index, end_column_index = (
            series_id.split("|")
        )
        updated_index_end = str(int(end_column_index) + column_delta)
        for series in self.series_list:
            if (
                series.series_id
                == f"{sheet_name}|{series_header}|{start_column_index}|{updated_index_end}"
            ):
                return series.series_id
        return series_id

    def process_function_node(self, node1, node2, n):
        modified_args = [
            self.apply_delta_to_range_node(arg, node2.args[i], n)
            for i, arg in enumerate(node1.args)
        ]
        modified_node = xlcalculator.ast_nodes.FunctionNode(node1.token)
        modified_node.args = modified_args
        return modified_node

    def process_operator_node(self, node1, node2, n):
        modified_left = (
            self.apply_delta_to_range_node(node1.left, node2.left, n)
            if node1.left
            else None
        )
        modified_right = (
            self.apply_delta_to_range_node(node1.right, node2.right, n)
            if node1.right
            else None
        )
        modified_node = xlcalculator.ast_nodes.OperatorNode(node1.token)
        modified_node.left = modified_left
        modified_node.right = modified_right
        return modified_node

    def get_nth_formula(self, n: int) -> xlcalculator.ast_nodes.ASTNode:
        return self.apply_delta_to_range_node(
            self.formula_1_ast_series, self.formula_2_ast_series, n=n
        )


class FormulaGenerator:
    """Creates instances of ASTGenerator given two formula_ast objects"""

    @staticmethod
    def get_ast_generator(
        formula_1_ast_series: xlcalculator.ast_nodes.ASTNode,
        formula_2_ast_series: xlcalculator.ast_nodes.ASTNode,
        series_list: List[Series],
    ) -> ASTGenerator:
        """Create an instance of ASTGenerator given two formula_ast objects and a series_list"""
        ast_generator = ASTGenerator(
            formula_1_ast_series, formula_2_ast_series, series_list
        )
        return ast_generator


# class ASTGenerator:
#     def __init__(
#         self,
#         formula_1_ast_series: xlcalculator.ast_nodes.ASTNode,
#         formula_2_ast_series: xlcalculator.ast_nodes.ASTNode,
#         series_list: List[Series],
#     ):
#         self.formula_1_ast_series = formula_1_ast_series
#         self.formula_2_ast_series = formula_2_ast_series
#         self.series_list = series_list

#     @staticmethod
#     def get_delta_between_nodes(node1_value: str, node2_value: str):

#         def extract_tuples(node_value: str):
#             return ast.literal_eval(node_value)

#         node1_tuple = extract_tuples(node1_value)
#         node2_tuple = extract_tuples(node2_value)

#         if node1_tuple[1] == (None, None) or node2_tuple[1] == (None, None):
#             return None
#         else:

#             node1_series_ids, node1_row_indexes = node1_tuple
#             node1_start_row_index, node1_end_row_index = node1_row_indexes

#             node1_start_column_index = int(node1_series_ids[0].split("|")[-1])
#             node1_end_column_index = int(node1_series_ids[-1].split("|")[-1])

#             node2_series_ids, node2_row_indexes = node2_tuple
#             node2_start_row_index, node2_end_row_index = node2_row_indexes

#             node2_start_column_index = int(node2_series_ids[0].split("|")[-1])
#             node2_end_column_index = int(node2_series_ids[-1].split("|")[-1])

#             start_row_index_delta = node2_start_row_index - node1_start_row_index
#             end_row_index_delta = node2_end_row_index - node1_end_row_index

#             start_column_index_delta = (
#                 node2_start_column_index - node1_start_column_index
#             )
#             end_column_index_delta = node2_end_column_index - node1_end_column_index

#             return (
#                 start_row_index_delta,
#                 end_row_index_delta,
#                 start_column_index_delta,
#                 end_column_index_delta,
#             ), (node1_start_row_index, node1_end_row_index)

#     def apply_delta_to_range_node(
#         self,
#         node1: xlcalculator.ast_nodes.ASTNode,
#         node2: xlcalculator.ast_nodes.ASTNode,
#         n: int,
#     ):
#         if isinstance(node1, xlcalculator.ast_nodes.RangeNode) and isinstance(
#             node2, xlcalculator.ast_nodes.RangeNode
#         ):
#             deltas_and_indexes = self.get_delta_between_nodes(
#                 node1.tvalue, node2.tvalue
#             )
#             if deltas_and_indexes:
#                 deltas, starting_indexes = deltas_and_indexes
#                 (
#                     start_row_index_delta,
#                     end_row_index_delta,
#                     start_column_index_delta,
#                     end_column_index_delta,
#                 ) = deltas
#                 (start_row_index, end_row_index) = starting_indexes

#                 series_ids = ast.literal_eval(node1.tvalue)[0]

#                 def add_column_delta_to_series_id(
#                     series_id: str, column_delta: int, series_list: List[Series]
#                 ):
#                     sheet_name, series_header, index_start, index_end = series_id.split(
#                         "|"
#                     )

#                     updated_index_end = str(int(index_end) + column_delta)

#                     # Find the series which starts with sheet_name and has index start and updated_index_end values
#                     series = next(
#                         (
#                             series
#                             for series in series_list
#                             if series.series_id.startswith(sheet_name)
#                             and series.series_id.endswith(updated_index_end)
#                             and series.series_id.split("|")[-2] == index_start
#                         ),
#                         None,
#                     )

#                     return series.series_id

#                 new_tvalue = str(
#                     (
#                         tuple(
#                             [
#                                 tuple(
#                                     [
#                                         add_column_delta_to_series_id(
#                                             series_id,
#                                             start_column_index_delta * (n - 1),
#                                             self.series_list,
#                                         )
#                                         for series_id in series_ids
#                                     ]
#                                 ),
#                                 (
#                                     (start_row_index + start_row_index_delta * (n - 1)),
#                                     (end_row_index + end_row_index_delta * (n - 1)),
#                                 ),
#                             ]
#                         )
#                     )
#                 )

#                 return xlcalculator.ast_nodes.RangeNode(
#                     xlcalculator.tokenizer.f_token(
#                         tvalue=new_tvalue, ttype="operand", tsubtype="range"
#                     )
#                 )
#             else:
#                 return node1

#         elif hasattr(node1, "args") and isinstance(
#             node1, xlcalculator.ast_nodes.FunctionNode
#         ):
#             modified_args = [
#                 self.apply_delta_to_range_node(arg, node2.args[i], n)
#                 for i, arg in enumerate(node1.args)
#             ]
#             modified_node = xlcalculator.ast_nodes.FunctionNode(node1.token)
#             modified_node.args = modified_args
#             return modified_node

#         elif hasattr(node1, "left") and isinstance(
#             node1, xlcalculator.ast_nodes.OperatorNode
#         ):
#             modified_left = (
#                 self.apply_delta_to_range_node(node1.left, node2.left, n)
#                 if node1.left
#                 else None
#             )
#             modified_right = (
#                 self.apply_delta_to_range_node(node1.right, node2.right, n)
#                 if node1.right
#                 else None
#             )
#             modified_node = xlcalculator.ast_nodes.OperatorNode(node1.token)
#             modified_node.left = modified_left
#             modified_node.right = modified_right
#             return modified_node

#         else:
#             return node1

#     def get_nth_formula(self, n: int) -> xlcalculator.ast_nodes.ASTNode:
#         return self.apply_delta_to_range_node(
#             self.formula_1_ast_series, self.formula_2_ast_series, n=n
#         )
