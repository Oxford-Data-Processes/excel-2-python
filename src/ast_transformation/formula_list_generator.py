import ast
import xlcalculator
import copy


class FormulaListGenerator:
    def __init__(self, formula_ast):
        self.formula_ast = formula_ast

    def adjust_indices(self, node, index_increment):
        if isinstance(node, xlcalculator.ast_nodes.FunctionNode):
            self.adjust_function_node_indices(node, index_increment)
        elif hasattr(node, "left") and hasattr(node, "right"):
            self.adjust_indices(node.left, index_increment)
            self.adjust_indices(node.right, index_increment)

    def adjust_function_node_indices(self, node, index_increment):
        new_args = []
        for arg in node.args:
            new_args.append(
                self.adjust_operand_node(arg, index_increment)
                if isinstance(arg, xlcalculator.ast_nodes.OperandNode)
                else arg
            )
        node.args = new_args

    def adjust_operand_node(self, arg, index_increment):
        parts = ast.literal_eval(arg.tvalue)
        series_tuple, indexes, deltas = parts
        updated_indexes = (indexes[0] + index_increment, indexes[1] + index_increment)
        updated_tvalue = f"(({repr(series_tuple[0])},), {updated_indexes}, {deltas})"
        return xlcalculator.ast_nodes.OperandNode(
            xlcalculator.tokenizer.f_token(
                tvalue=updated_tvalue, ttype="operand", tsubtype="text"
            )
        )

    def generate_formulas_list(self, start_index: int, end_index: int):
        return [
            self.generate_single_formula(i) for i in range(start_index, end_index + 1)
        ]

    def generate_single_formula(self, index_increment):
        new_ast = copy.deepcopy(self.formula_ast)
        self.adjust_indices(new_ast, index_increment)
        return new_ast
