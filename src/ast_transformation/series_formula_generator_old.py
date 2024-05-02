from ast_transformation.cell_range_implementer import CellRangeImplementer
from ast_transformation.formula_checker import FormulaChecker
from ast_transformation.formula_generator_old import FormulaGenerator


class SeriesFormulaGenerator:

    @staticmethod
    def adjust_formulas(formulas):
        formula_1, formula_2 = formulas
        if formula_1 is not None and formula_2 is None:
            formula_2 = formula_1
        return formula_1, formula_2

    @staticmethod
    def process_series_formulas(
        series,
        formula_1_ast_series,
        formula_2_ast_series,
        series_mapping,
        series_dict,
        series_list_new,
    ):

        sheet_name = series.worksheet.sheet_name
        series_list_within_sheet = series_dict.get(sheet_name)
        ast_generator = FormulaGenerator.get_ast_generator(
            formula_1_ast_series, formula_2_ast_series, series_list_within_sheet
        )

        formulas_are_correct, formula_1_ast_new, formula_2_ast_new = (
            FormulaChecker.check_formulas(ast_generator)
        )
        if not formulas_are_correct:
            raise Exception("Formulas are not correct")

        SeriesFormulaGenerator.update_series_with_new_asts(
            series, formula_1_ast_new, formula_2_ast_new, series_dict, series_list_new
        )

        return ast_generator

    @staticmethod
    def update_series_with_new_asts(
        series, formula_1_ast_new, formula_2_ast_new, series_dict, series_list_new
    ):
        cell_range_implementer = CellRangeImplementer(series_dict)
        formula_1_ast_new_cell_ranges = cell_range_implementer.update_ast(
            formula_1_ast_new
        )
        formula_2_ast_new_cell_ranges = cell_range_implementer.update_ast(
            formula_2_ast_new
        )

        series.formulas = [
            f"={formula_1_ast_new_cell_ranges}",
            f"={formula_2_ast_new_cell_ranges}",
        ]
        series_list_new.append(series)
