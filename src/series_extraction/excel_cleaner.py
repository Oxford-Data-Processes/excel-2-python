import re
import openpyxl
from objects import ExcelFile


class ExcelCleaner:

    @staticmethod
    def remove_quotes(formula: str) -> str:
        return re.sub(r"'(.*?)'!", r"\1!", formula)

    @staticmethod
    def remove_dollar_sign_from_excel_formula(formula: str) -> str:
        return re.sub(r"\$", "", formula)

    @staticmethod
    def clean_excel(excel_reduced: ExcelFile) -> ExcelFile:
        for sheet in excel_reduced.workbook_with_formulas.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        formula = cell.value

                        if isinstance(formula, openpyxl.worksheet.formula.ArrayFormula):
                            formula = formula.text

                        formula = ExcelCleaner.remove_quotes(formula)
                        formula = ExcelCleaner.remove_dollar_sign_from_excel_formula(
                            formula
                        )
                        cell.value = formula
        return excel_reduced
