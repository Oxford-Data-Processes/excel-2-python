from objects import ExcelFile, Worksheet, Table
from typing import Dict, List


class ExcelCompatibilityChecker:
    @staticmethod
    def check_file(
        excel_raw: ExcelFile,
        excel_reduced: ExcelFile,
        extracted_tables: Dict[Worksheet, List[Table]],
    ) -> bool:
        for worksheet, tables in extracted_tables.items():
            for table in tables:
                raw_ws = excel_raw.workbook_with_values[worksheet.sheet_name]
                reduced_ws = excel_reduced.workbook_with_values[worksheet.sheet_name]

                raw_headers = [
                    cell.value for cell in raw_ws[table.range.start_cell.row]
                ]
                reduced_headers = [
                    cell.value for cell in reduced_ws[table.range.start_cell.row]
                ]

                if set(raw_headers) != set(reduced_headers):

                    return False

        return True
