import ast


class FormulaEvaluator:

    def __init__(self, series_dict):
        self.local_env = {"SUM": self.SUM}
        self.series_dict = series_dict

    def get_values_from_series(self, series_tuple):
        series_ids, indexes, _ = series_tuple
        series_values = self.series_dict[series_ids[0]]
        start_index, end_index = indexes
        return series_values[start_index:end_index]

    def SUM(self, series_tuple_string: str):
        series_tuple = ast.literal_eval(series_tuple_string)
        return sum(self.get_values_from_series(series_tuple))

    def evaluate_sum(self, sum_formula: str):
        tree = ast.parse(sum_formula, mode="eval")
        compiled = compile(tree, filename="<ast>", mode="eval")
        result = eval(compiled, {"__builtins__": {}}, self.local_env)
        return result
