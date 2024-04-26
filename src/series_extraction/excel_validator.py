import re
from objects import ExcelFile


class ExcelValidator:
    @staticmethod
    def validate_excel(excel_reduced: ExcelFile) -> bool:
        """Check if the workbook's sheet titles do not contain spaces, returning True if valid."""
        is_valid = True
        for sheet in excel_reduced.workbook_with_formulas.worksheets:
            if " " in sheet.title:
                is_valid = False
                break
        return is_valid
