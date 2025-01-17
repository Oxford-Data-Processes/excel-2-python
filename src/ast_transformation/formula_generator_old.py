import xlcalculator
import ast
from typing import List, Tuple, Optional, Callable
from objects import Series, SeriesId


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
    def get_delta_between_nodes(node1_value: str, node2_value: str):
        node1_tuple = ast.literal_eval(node1_value)
        node2_tuple = ast.literal_eval(node2_value)

        if node1_tuple[1] == (None, None) or node2_tuple[1] == (None, None):
            return None
        else:
            return DeltaCalculator.calculate_deltas(node1_tuple, node2_tuple)

    @staticmethod
    def load_series_ids(series_ids_strings: List[str]) -> List[SeriesId]:
        return [
            SeriesIdLoader.load_series_id_from_string(sid) for sid in series_ids_strings
        ]

    @staticmethod
    def calculate_index_deltas(
        indexes1: Tuple[int, int], indexes2: Tuple[int, int]
    ) -> Tuple[int, int]:
        return tuple(x - y for x, y in zip(indexes1, indexes2))

    @staticmethod
    def calculate_series_id_index_deltas(
        series_ids1: List[SeriesId], series_ids2: List[SeriesId]
    ) -> Tuple[int, int, int, int]:
        start_row_index_delta = (
            series_ids2[0].series_header_cell_row
            - series_ids1[0].series_header_cell_row
        )
        end_row_index_delta = (
            series_ids2[-1].series_header_cell_row
            - series_ids1[-1].series_header_cell_row
        )
        start_column_index_delta = (
            series_ids2[0].series_header_cell_column
            - series_ids1[0].series_header_cell_column
        )
        end_column_index_delta = (
            series_ids2[-1].series_header_cell_column
            - series_ids1[-1].series_header_cell_column
        )
        return (
            start_row_index_delta,
            end_row_index_delta,
            start_column_index_delta,
            end_column_index_delta,
        )

    @staticmethod
    def calculate_deltas(
        node1_tuple: Tuple[List[str], Tuple[int, int]],
        node2_tuple: Tuple[List[str], Tuple[int, int]],
    ) -> SeriesRangeDelta:
        node1_series_ids_strings, node1_row_indexes = node1_tuple
        node2_series_ids_strings, node2_row_indexes = node2_tuple

        # Load SeriesId objects from strings
        node1_series_ids = DeltaCalculator.load_series_ids(node1_series_ids_strings)
        node2_series_ids = DeltaCalculator.load_series_ids(node2_series_ids_strings)

        # Calculate row index deltas
        row_index_deltas = DeltaCalculator.calculate_index_deltas(
            node2_row_indexes, node1_row_indexes
        )

        # Calculate series ID index deltas
        series_id_deltas = DeltaCalculator.calculate_series_id_index_deltas(
            node1_series_ids, node2_series_ids
        )

        # Construct and return the SeriesRangeDelta object
        return SeriesRangeDelta(
            *row_index_deltas, *series_id_deltas, *node1_row_indexes
        )


class NodeProcessor:
    def __init__(self, series_list: List[Series]):
        self.series_list = series_list
        self.series_updater = SeriesUpdater(series_list)

    def process_range_node(
        self,
        node1: xlcalculator.ast_nodes.RangeNode,
        node2: xlcalculator.ast_nodes.RangeNode,
        n: int,
    ) -> xlcalculator.ast_nodes.RangeNode:
        series_range_delta = DeltaCalculator.get_delta_between_nodes(
            node1.tvalue, node2.tvalue
        )
        if series_range_delta:
            return self.apply_delta_to_node(node1, series_range_delta, n)
        return node1

    def process_function_node(
        self,
        node1: xlcalculator.ast_nodes.FunctionNode,
        node2: xlcalculator.ast_nodes.FunctionNode,
        n: int,
    ) -> xlcalculator.ast_nodes.FunctionNode:
        modified_args = [
            self.apply_delta_to_range_node(arg, node2.args[i], n)
            for i, arg in enumerate(node1.args)
        ]
        modified_node = xlcalculator.ast_nodes.FunctionNode(node1.token)
        modified_node.args = modified_args
        return modified_node

    def process_operator_node(
        self,
        node1: xlcalculator.ast_nodes.OperatorNode,
        node2: xlcalculator.ast_nodes.OperatorNode,
        n: int,
    ) -> xlcalculator.ast_nodes.OperatorNode:
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

    def apply_delta_to_node(
        self,
        node: xlcalculator.ast_nodes.ASTNode,
        series_range_delta: SeriesRangeDelta,
        n: int,
    ) -> xlcalculator.ast_nodes.ASTNode:
        deltas = self.extract_deltas_from_range(series_range_delta)
        return self.update_range_node(node, *deltas, n)

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

    def extract_deltas_from_range(
        self, series_range_delta: SeriesRangeDelta
    ) -> Tuple[int, int, int, int, int, int, int, int]:
        return (
            series_range_delta.start_row_index_delta,
            series_range_delta.end_row_index_delta,
            series_range_delta.series_id_start_row_index_delta,
            series_range_delta.series_id_end_row_index_delta,
            series_range_delta.series_id_start_column_index_delta,
            series_range_delta.series_id_end_column_index_delta,
            series_range_delta.start_row_index,
            series_range_delta.end_row_index,
        )

    def update_range_node(
        self,
        node1: xlcalculator.ast_nodes.RangeNode,
        start_row_index_delta: int,
        end_row_index_delta: int,
        series_id_start_row_index_delta: int,
        series_id_end_row_index_delta: int,
        series_id_start_column_index_delta: int,
        series_id_end_column_index_delta: int,
        start_row_index: int,
        end_row_index: int,
        n: int,
    ) -> xlcalculator.ast_nodes.RangeNode:

        series_ids_string = ast.literal_eval(node1.tvalue)[0]
        series_ids = [
            SeriesIdLoader.load_series_id_from_string(sid) for sid in series_ids_string
        ]

        new_series_ids = [
            str(
                self.series_updater.add_column_delta_to_series_id(
                    sid,
                    series_id_start_row_index_delta * (n - 1),
                    series_id_start_column_index_delta * (n - 1),
                )
            )
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


class SeriesUpdater:
    def __init__(self, series_list):
        self.series_list = series_list

    def update_series_header_indexes(
        self, series_id: SeriesId, row_delta: int, column_delta: int
    ) -> Tuple[int, int]:
        """Update the series header indexes by adding the respective deltas."""
        return (
            series_id.series_header_cell_column + column_delta,
            series_id.series_header_cell_row + row_delta,
        )

    def find_series(self, condition: Callable[[Series], bool]) -> Optional[SeriesId]:
        """Generalized method to find a series based on a provided condition."""
        for series in self.series_list:
            if condition(series):
                return series.series_id
        return None

    def add_column_delta_to_series_id(
        self,
        series_id: SeriesId,
        series_id_start_row_index_delta: int,
        series_id_start_column_index_delta: int,
    ):
        """Apply column and row deltas to the series ID and find the matching series ID."""
        updated_column, updated_row = self.update_series_header_indexes(
            series_id,
            series_id_start_row_index_delta,
            series_id_start_column_index_delta,
        )

        if (
            series_id_start_column_index_delta > 0
            and series_id_start_row_index_delta == 0
        ):
            return (
                self.find_series(
                    lambda s: s.series_id.sheet_name == series_id.sheet_name
                    and s.series_id.series_header_cell_column == updated_column
                )
                or series_id
            )

        elif (
            series_id_start_column_index_delta == 0
            and series_id_start_row_index_delta > 0
        ):
            return (
                self.find_series(
                    lambda s: s.series_id.sheet_name == series_id.sheet_name
                    and s.series_id.series_header_cell_row == updated_row
                )
                or series_id
            )

        elif (
            series_id_start_column_index_delta > 0
            and series_id_start_row_index_delta > 0
        ):
            return (
                self.find_series(
                    lambda s: s.series_id.sheet_name == series_id.sheet_name
                    and s.series_id.series_header_cell_row == updated_row
                    and s.series_id.series_header_cell_column == updated_column
                )
                or series_id
            )

        else:
            return series_id


class ASTGenerator:
    def __init__(
        self,
        formula_1_ast_series: xlcalculator.ast_nodes.ASTNode,
        formula_2_ast_series: xlcalculator.ast_nodes.ASTNode,
        series_list: List[Series],
    ):
        self.formula_1_ast_series = formula_1_ast_series
        self.formula_2_ast_series = formula_2_ast_series
        self.node_processor = NodeProcessor(series_list)

    def get_nth_formula(self, n: int) -> xlcalculator.ast_nodes.ASTNode:
        return self.node_processor.apply_delta_to_range_node(
            self.formula_1_ast_series, self.formula_2_ast_series, n=n
        )


class FormulaGenerator:
    """Creates generic formula given two formula_ast objects"""

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
