import re
from objects import ExcelFile


class ExcelValidator:
    @staticmethod
    def validate_excel(excel_reduced: ExcelFile) -> bool:
        is_valid = True
        for sheet in excel_reduced.workbook_with_formulas.worksheets:
            if " " in sheet.title:
                is_valid = False
                break
        return is_valid
