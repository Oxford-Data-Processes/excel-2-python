import ast
import xlcalculator
import copy
from xlcalculator.ast_nodes import RangeNode, FunctionNode, OperatorNode
from xlcalculator.tokenizer import f_token


class FormulaListGenerator:
    def __init__(self, formula_ast, series_values_dict):
        self.formula_ast = formula_ast
        self.series_values_dict = series_values_dict

    def update_ast(self, node, index_increment):
        if isinstance(node, RangeNode):
            return self.replace_range_node(node, index_increment)
        elif isinstance(node, FunctionNode):
            return self.replace_function_node(node, index_increment)
        elif isinstance(node, OperatorNode):
            return self.replace_operator_node(node, index_increment)
        return node

    def replace_range_node(self, node, index_increment):
        parts = ast.literal_eval(node.tvalue)
        series_tuple, indexes, deltas = parts
        start_new_index, end_new_index = (
            indexes[0] + index_increment,
            indexes[1] + index_increment,
        )

        array_values = [
            self.series_values_dict.get(series_id)[start_new_index : end_new_index + 1]
            for series_id in series_tuple
        ]

        array_row_nodes = []
        for row in array_values:
            row_node = FunctionNode(
                f_token(tvalue="ARRAYROW", ttype="function", tsubtype="")
            )
            row_node_args = []
            for value in row:
                if isinstance(value, str):
                    value = f'"{value}"'
                row_node_args.append(value)

            row_node.args = row_node_args
            array_row_nodes.append(row_node)

        array_node = FunctionNode(
            f_token(tvalue="ARRAY", ttype="function", tsubtype="")
        )
        array_node.args = array_row_nodes

        return array_node

    def replace_function_node(self, node, index_increment):
        modified_args = [self.update_ast(arg, index_increment) for arg in node.args]
        modified_function_node = FunctionNode(node.token)
        modified_function_node.args = modified_args
        return modified_function_node

    def replace_operator_node(self, node, index_increment):
        modified_left = (
            self.update_ast(node.left, index_increment) if node.left else None
        )
        modified_right = (
            self.update_ast(node.right, index_increment) if node.right else None
        )
        modified_operator_node = OperatorNode(node.token)
        modified_operator_node.left = modified_left
        modified_operator_node.right = modified_right
        return modified_operator_node

    def adjust_indices(self, ast, index_increment):
        return self.update_ast(ast, index_increment)

    def generate_formula_list(self, start_index: int, end_index: int):
        return [
            self.generate_single_formula(i) for i in range(start_index, end_index + 1)
        ]

    def generate_single_formula(self, index_increment):
        new_ast = copy.deepcopy(self.formula_ast)
        new_ast = self.adjust_indices(new_ast, index_increment)
        return new_ast
