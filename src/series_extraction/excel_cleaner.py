import re
import openpyxl
from openpyxl.cell.cell import Cell
from objects import ExcelFile


class ExcelCleaner:

    @staticmethod
    def remove_quotes(formula: str) -> str:
        """Remove single quotes from cell references in Excel formulas."""
        return re.sub(r"'(.*?)'!", r"\1!", formula)

    @staticmethod
    def remove_dollar_sign_from_excel_formula(formula: str) -> str:
        """Remove dollar signs from Excel formulas to convert absolute references to relative."""
        return re.sub(r"\$", "", formula)

    @staticmethod
    def get_cell_formula(cell: Cell) -> str:
        """Extract the formula from a cell, accounting for both regular and array formulas."""
        if isinstance(cell.value, openpyxl.worksheet.formula.ArrayFormula):
            return cell.value.text
        return cell.value

    @staticmethod
    def clean_formula(formula: str) -> str:
        """Clean an Excel formula by removing quotes and dollar signs."""
        formula = ExcelCleaner.remove_quotes(formula)
        formula = ExcelCleaner.remove_dollar_sign_from_excel_formula(formula)
        return formula

    @staticmethod
    def clean_excel(excel_reduced: ExcelFile) -> ExcelFile:
        """Clean all formulas in an Excel file by removing quotes and dollar signs."""
        for sheet in excel_reduced.workbook_with_formulas.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        formula = ExcelCleaner.get_cell_formula(cell)
                        cleaned_formula = ExcelCleaner.clean_formula(formula)
                        cell.value = cleaned_formula
        return excel_reduced
