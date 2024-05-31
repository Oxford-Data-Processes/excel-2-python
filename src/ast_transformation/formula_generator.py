from formulas.tokens.operand import Range
from excel_utils import ExcelUtils


class FormulaGenerator:

    @staticmethod
    def get_delta_between_coordinates(coordinate1, coordinate2):

        coordinate1_cell_range = coordinate1.split("!")[-1]
        coordinate2_cell_range = coordinate2.split("!")[-1]
        (
            cell_start_column2,
            cell_start_row2,
            cell_end_column2,
            cell_end_row2,
            is_column_range,
        ) = ExcelUtils.get_coordinates_from_range(coordinate2_cell_range)
        (
            cell_start_column1,
            cell_start_row1,
            cell_end_column1,
            cell_end_row1,
            is_column_range,
        ) = ExcelUtils.get_coordinates_from_range(coordinate1_cell_range)

        cell_start_column_delta = cell_start_column2 - cell_start_column1
        cell_start_row_delta = cell_start_row2 - cell_start_row1
        cell_end_column_delta = cell_end_column2 - cell_end_column1
        cell_end_row_delta = cell_end_row2 - cell_end_row1

        return (
            cell_start_column_delta,
            cell_start_row_delta,
            cell_end_column_delta,
            cell_end_row_delta,
        )

    @staticmethod
    def generate_generic_formula(formula_1_elements_list, formula_2_elements_list):

        delta_dictionary = {}

        for index, (element1, element2) in enumerate(
            zip(formula_1_elements_list, formula_2_elements_list)
        ):

            if isinstance(element1, Range) and isinstance(element2, Range):

                delta = FormulaGenerator.get_delta_between_coordinates(
                    element1.name, element2.name
                )
                delta_dictionary[index] = delta

        return delta_dictionary
