import xlcalculator
import ast
import copy


class ASTGenerator:
    def __init__(
        self,
        formula_1_ast_series: xlcalculator.ast_nodes.ASTNode,
        formula_2_ast_series: xlcalculator.ast_nodes.ASTNode,
    ):
        self.formula_1_ast_series = formula_1_ast_series
        self.formula_2_ast_series = formula_2_ast_series

    @staticmethod
    def get_delta_between_nodes(node1_value: str, node2_value: str):

        def extract_tuples(node_value: str):
            return ast.literal_eval(node_value)

        node1_tuple = extract_tuples(node1_value)
        node2_tuple = extract_tuples(node2_value)

        if node1_tuple[1] == (None, None) or node2_tuple[1] == (None, None):
            return None
        else:

            _, node1_row_indexes = node1_tuple
            node1_start_row_index, node1_end_row_index = node1_row_indexes
            _, node2_row_indexes = node2_tuple
            node2_start_row_index, node2_end_row_index = node2_row_indexes

        start_row_index_delta = node2_start_row_index - node1_start_row_index
        end_row_index_delta = node2_end_row_index - node1_end_row_index

        return (
            start_row_index_delta,
            end_row_index_delta,
        ), (node1_start_row_index, node1_end_row_index)

    def apply_delta_to_range_node(
        self,
        node1: xlcalculator.ast_nodes.ASTNode,
        node2: xlcalculator.ast_nodes.ASTNode,
        n: int,
    ):
        if isinstance(node1, xlcalculator.ast_nodes.RangeNode) and isinstance(
            node2, xlcalculator.ast_nodes.RangeNode
        ):
            deltas, starting_indexes = self.get_delta_between_nodes(
                node1.tvalue, node2.tvalue
            )
            if deltas:
                (start_row_index_delta, end_row_index_delta) = deltas
                (start_row_index, end_row_index) = starting_indexes

                series_ids = ast.literal_eval(node1.tvalue)[0]
                new_tvalue = str(
                    (
                        tuple(
                            [
                                series_ids,
                                (
                                    start_row_index + start_row_index_delta * (n - 1),
                                    end_row_index + end_row_index_delta * (n - 1),
                                ),
                            ]
                        )
                    )
                )

                return xlcalculator.ast_nodes.RangeNode(
                    xlcalculator.tokenizer.f_token(
                        tvalue=new_tvalue, ttype="operand", tsubtype="range"
                    )
                )
            else:
                return node1

        elif hasattr(node1, "args") and isinstance(
            node1, xlcalculator.ast_nodes.FunctionNode
        ):
            modified_args = [
                self.apply_delta_to_range_node(arg, node2.args[i], n)
                for i, arg in enumerate(node1.args)
            ]
            modified_node = xlcalculator.ast_nodes.FunctionNode(node1.token)
            modified_node.args = modified_args
            return modified_node

        elif hasattr(node1, "left") and isinstance(
            node1, xlcalculator.ast_nodes.OperatorNode
        ):
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

        else:
            return node1

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
    ) -> ASTGenerator:
        """Create an instance of ASTGenerator given two formula_ast objects"""
        ast_generator = ASTGenerator(formula_1_ast_series, formula_2_ast_series)
        return ast_generator
