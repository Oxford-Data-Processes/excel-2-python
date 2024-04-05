import re
from objects import ExcelFile


class ExcelValidator:
    @staticmethod
    def validate_excel(excel_reduced: ExcelFile) -> bool:
        """False if there is a space in one of the sheet names, True otherwise"""
        is_valid = True
        for sheet in excel_reduced.workbook.worksheets:
            if " " in sheet.title:
                is_valid = False
                break
        return is_valid
