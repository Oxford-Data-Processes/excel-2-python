import ast
import numpy as np
from numpy.typing import NDArray
from typing import Union


class FormulaEvaluator:
    def __init__(self, series_dict):
        self.local_env = {"SUM": self.SUM}
        self.series_dict = series_dict

    def get_values_from_series(self, series_tuple) -> NDArray[np.float64]:
        series_ids, indexes, _ = series_tuple
        values = np.array([], dtype=np.float64)  # Ensure the array is of type float64
        for series_id in series_ids:
            series_values = self.series_dict[series_id]
            start_index, end_index = indexes
            values = np.concatenate(
                (values, series_values[start_index : end_index + 1])
            )
        return values

    def SUM(self, series_tuple_or_string: Union[str, tuple]):
        if isinstance(series_tuple_or_string, str):
            series_tuple_string = series_tuple_or_string
        else:
            series_tuple_string = str(series_tuple_or_string)

        series_tuple = ast.literal_eval(series_tuple_string)
        return sum(self.get_values_from_series(series_tuple))

    def evaluate_formula(self, formula: str):
        tree = ast.parse(formula, mode="eval")
        compiled = compile(tree, filename="<ast>", mode="eval")
        result = eval(compiled, {"__builtins__": {}}, self.local_env)
        return result
