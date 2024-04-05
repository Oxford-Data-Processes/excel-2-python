import openpyxl
from objects import ExcelFile


class ExcelLoader:
    @staticmethod
    def load_file(file_path: str) -> ExcelFile:
        workbook = openpyxl.load_workbook(file_path)
        return ExcelFile(file_path=file_path, workbook=workbook)
