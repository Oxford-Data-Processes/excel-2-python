import xlcalculator
import pandas as pd
import ast
import numpy as np

from ast_transformation.formula_generator_old import SeriesIdLoader


class FormulaEvaluator:

    def __init__(self, formula_ast, series_dict):
        self.series_dict = series_dict

        range_nodes_string_list = FormulaEvaluator.extract_series_id_string_list(
            formula_ast
        )
        series_id_string_list = FormulaEvaluator.get_series_id_string_list(
            range_nodes_string_list
        )

        self.df_dict = {}
        for series_id_string in series_id_string_list:
            df = self.get_dataframe_from_series_id_string(
                series_id_string, self.series_dict
            )
            self.df_dict[series_id_string] = df

    @staticmethod
    def extract_series_id_string_list(ast):

        series_id_string_list = []

        def replace_range_node(node):
            return node

        def replace_function_node(node):
            modified_args = [traverse_ast(arg) for arg in node.args]
            modified_function_node = xlcalculator.ast_nodes.FunctionNode(node.token)
            modified_function_node.args = modified_args
            return modified_function_node

        def replace_operator_node(node):
            modified_left = traverse_ast(node.left) if node.left else None
            modified_right = traverse_ast(node.right) if node.right else None
            modified_operator_node = xlcalculator.ast_nodes.OperatorNode(node.token)
            modified_operator_node.left = modified_left
            modified_operator_node.right = modified_right
            return modified_operator_node

        def traverse_ast(node):
            if isinstance(node, xlcalculator.ast_nodes.RangeNode):
                series_id_string_list.append(node.tvalue)
                return replace_range_node(node)
            elif isinstance(node, xlcalculator.ast_nodes.FunctionNode):
                return replace_function_node(node)
            elif isinstance(node, xlcalculator.ast_nodes.OperatorNode):
                return replace_operator_node(node)
            elif isinstance(node, list):
                return [traverse_ast(item) for item in node]
            elif hasattr(node, "children"):
                traverse_ast(node.children)
            return node

        traverse_ast(ast)
        return series_id_string_list

    @staticmethod
    def get_dataframe_from_series_id_string(series_id_string, series_dict):
        # Assume series_dict and SeriesIdLoader are predefined somewhere in the project
        df_list = []
        series_id = SeriesIdLoader.load_series_id_from_string(series_id_string)
        sheet_name = series_id.sheet_name
        for series in series_dict[sheet_name]:
            if series.series_id == series_id:
                df_list.append(
                    pd.DataFrame(
                        data=series.values, columns=[series.series_id.series_header]
                    )
                )
        return pd.concat(df_list, axis=1)

    def fetch_df(self, identifier, index_range):
        df = self.df_dict[identifier]
        start, end = index_range
        return df.iloc[start : end + 1]

    def IF(self, *args):
        if len(args) != 3:
            raise ValueError("IF function expects exactly three arguments.")
        return xlcalculator.xlfunctions.logical.IF(*args)

    def AND(self, *args):
        return xlcalculator.xlfunctions.logical.AND(*args)

    def ROUND(self, *args):

        value, decimal_points = args
        if isinstance(value, float):
            return xlcalculator.xlfunctions.math.ROUND(value, decimal_points)
        else:
            identifiers, index_range = value

            series_df = self.fetch_df(identifiers[0], index_range)
            number = float(series_df.iloc[0, 0])

            return xlcalculator.xlfunctions.math.ROUND(number, decimal_points)

    def ROUNDDOWN(self, *args):

        value, decimal_points = args
        if isinstance(value, float):
            return xlcalculator.xlfunctions.math.ROUNDDOWN(value, decimal_points)
        else:
            identifiers, index_range = value

            series_df = self.fetch_df(identifiers[0], index_range)
            number = float(series_df.iloc[0, 0])

            return xlcalculator.xlfunctions.math.ROUNDDOWN(number, decimal_points)

    def AVERAGE(self, args):
        identifiers, index_range = args

        series_list = [
            self.fetch_df(identifier, index_range) for identifier in identifiers
        ]

        numbers = [
            item
            for sublist in [
                series.select_dtypes(include=[np.number]).values.flatten()
                for series in series_list
            ]
            for item in sublist
        ]

        return xlcalculator.xlfunctions.statistics.AVERAGE(numbers)

    def VLOOKUP(self, *args):
        value, table, column, exact_match = args
        identifiers, index_range = value
        series_df = self.fetch_df(identifiers[0], index_range)
        lookup_value = series_df.iloc[0, 0]
        table_identifiers, table_index_range = table
        table_series_list = [
            self.fetch_df(table_identifier, (0, 1))
            for table_identifier in table_identifiers
        ]

        return xlcalculator.xlfunctions.lookup.VLOOKUP(
            lookup_value, table_series_list, column, exact_match
        )

    def SUM(self, *args):
        identifiers, index_range = args[0]
        series_list = [
            self.fetch_df(identifier, index_range) for identifier in identifiers
        ]
        numbers = [
            item
            for sublist in [
                series.select_dtypes(include=[np.number]).values.flatten()
                for series in series_list
            ]
            for item in sublist
        ]
        return xlcalculator.xlfunctions.math.SUM(numbers)

    def evaluate_formula(self, formula):
        formula = str(formula).replace("=", "==")
        tree = ast.parse(formula, mode="eval")
        local_env = {
            "AVERAGE": self.AVERAGE,
            "VLOOKUP": self.VLOOKUP,
            "IF": self.IF,
            "AND": self.AND,
            "ROUND": self.ROUND,
            "ROUNDDOWN": self.ROUNDDOWN,
            "SUM": self.SUM,
        }
        compiled = compile(tree, filename="<ast>", mode="eval")
        result = eval(compiled, {"__builtins__": {}}, local_env)
        if isinstance(result, xlcalculator.xlfunctions.func_xltypes.Number):
            return float(result)
        if isinstance(result, xlcalculator.xlfunctions.xlerrors.ValueExcelError):
            return "#N/A"
        # print("result type")
        # print(type(result))
        return result

    @staticmethod
    def get_series_id_string_list(range_nodes_string_list):
        series_id_string_list = []

        for item in range_nodes_string_list:
            item = ast.literal_eval(item)[0]
            for i in item:
                series_id_string_list.append(i)
        return series_id_string_list

    def get_series_from_id(self, series_id):
        for series in self.series_dict[series_id.sheet_name]:
            if series.series_id == series_id:
                return series

    def calculate_series_values(self, ast_generator, index_start, index_end):
        values = []
        for index in range(index_start, index_end + 1):
            formula = ast_generator.get_nth_formula(n=index)
            formula_evaluator = FormulaEvaluator(formula, self.series_dict)
            value = formula_evaluator.evaluate_formula(formula)
            values.append(value)
        return values
