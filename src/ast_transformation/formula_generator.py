import xlcalculator
import re


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
        def check_string(s):
            pattern = r"_\d_\d$"
            return bool(re.search(pattern, s))

        print(node1_value, node2_value)

        if node1_value == node2_value:
            return None
        elif (
            check_string(node1_value)
            and check_string(node2_value)
            and node1_value[:-4] == node2_value[:-4]
        ):
            node_1_start_index = int(node1_value[-3])
            node_1_end_index = int(node1_value[-1])
            node_2_start_index = int(node2_value[-3])
            node_2_end_index = int(node2_value[-1])
            if (
                node_1_start_index != node_2_start_index
                or node_1_end_index != node_2_end_index
            ):
                start_index_delta = node_2_start_index - node_1_start_index
                end_index_delta = node_2_end_index - node_1_end_index
                return (start_index_delta, end_index_delta)

    def apply_deltas_to_range_nodes(
        self,
        node1: xlcalculator.ast_nodes.ASTNode,
        node2: xlcalculator.ast_nodes.ASTNode,
    ):
        if isinstance(node1, xlcalculator.ast_nodes.RangeNode) and isinstance(
            node2, xlcalculator.ast_nodes.RangeNode
        ):
            deltas = self.get_delta_between_nodes(node1.tvalue, node2.tvalue)
            if deltas:
                start_index_delta, end_index_delta = deltas
                new_tvalue = f"{node1.tvalue[:-4]}_{start_index_delta}_{end_index_delta}_{node1.tvalue[-3]}"
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
                self.apply_deltas_to_range_nodes(arg, node2.args[i])
                for i, arg in enumerate(node1.args)
            ]
            modified_node = xlcalculator.ast_nodes.FunctionNode(node1.token)
            modified_node.args = modified_args
            return modified_node

        elif hasattr(node1, "left") and isinstance(
            node1, xlcalculator.ast_nodes.OperatorNode
        ):
            modified_left = (
                self.apply_deltas_to_range_nodes(node1.left, node2.left)
                if node1.left
                else None
            )
            modified_right = (
                self.apply_deltas_to_range_nodes(node1.right, node2.right)
                if node1.right
                else None
            )
            modified_node = xlcalculator.ast_nodes.OperatorNode(node1.token)
            modified_node.left = modified_left
            modified_node.right = modified_right
            return modified_node

        else:
            return node1

    def get_ast_with_deltas(self) -> xlcalculator.ast_nodes.ASTNode:
        return self.apply_deltas_to_range_nodes(
            self.formula_1_ast_series, self.formula_2_ast_series
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
