import copy
import xlcalculator
import xlcalculator.ast_nodes

from series_extraction.series_mapper import SeriesMapper
from ast_transformation.formula_list_generator import FormulaListGenerator
from ast_transformation.formula_evaluator import FormulaEvaluator

from typing import List, Dict, Any

from objects import SeriesId, Series


class PipelineBuilder:

    @staticmethod
    def get_evaluated_results_from_formula_ast(
        series_id: SeriesId,
        formula_ast: xlcalculator.ast_nodes.ASTNode,
        series_values_dict: Dict[str, List[Any]],
    ):
        formula_list_generator = FormulaListGenerator(formula_ast, series_values_dict)
        values_length = len(series_values_dict[str(series_id)])
        formula_list = formula_list_generator.generate_formula_list(
            0, values_length - 1
        )
        formula_strings = [str(formula) for formula in formula_list]

        formula_evaluator = FormulaEvaluator()

        results = [
            formula_evaluator.evaluate_formula(f"={formula_string}")
            for formula_string in formula_strings
        ]

        return results

    @staticmethod
    def create_series_list(
        sorted_dag: List[SeriesId],
        generic_formula_dictionary: Dict[SeriesId, xlcalculator.ast_nodes.ASTNode],
        series_dict: Dict[str, List[Series]],
        series_values_dict_raw: Dict[str, List[Any]],
        series_list_with_values: List[Series],
    ):

        sorted_dag = copy.deepcopy(sorted_dag)
        generic_formula_dictionary = copy.deepcopy(generic_formula_dictionary)
        series_dict = copy.deepcopy(series_dict)
        series_values_dict_raw = copy.deepcopy(series_values_dict_raw)
        series_list_with_values = copy.deepcopy(series_list_with_values)

        series_list_new_raw = []
        for series_id in sorted_dag:
            formula_ast = generic_formula_dictionary.get(series_id)

            series = SeriesMapper.get_series_from_series_id(series_id, series_dict)
            if formula_ast:
                values = PipelineBuilder.get_evaluated_results_from_formula_ast(
                    series_id, formula_ast, series_values_dict_raw
                )
                series.values = values
                series.series_length = len(values)
                series_list_new_raw.append(series)

        series_list_with_values_raw = []

        for series in series_list_with_values:
            series_id = series.series_id
            values = series_values_dict_raw[str(series_id)]
            series.values = values
            series.series_length = len(values)
            series_list_with_values_raw.append(series)

        series_list_updated_raw = series_list_new_raw + series_list_with_values_raw

        return series_list_updated_raw
