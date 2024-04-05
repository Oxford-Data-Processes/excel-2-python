import openpyxl
from objects import ExcelFile


class ExcelLoader:
    @staticmethod
    def load_file(file_path: str) -> ExcelFile:
        workbook_with_formulas = openpyxl.load_workbook(file_path, data_only=False)
        workbook_with_values = openpyxl.load_workbook(file_path, data_only=True)
        return ExcelFile(
            file_path=file_path,
            workbook_with_formulas=workbook_with_formulas,
            workbook_with_values=workbook_with_values,
        )
