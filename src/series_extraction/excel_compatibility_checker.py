from objects import ExcelFile, Worksheet, Table, HeaderLocation
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

                if table.header_location == HeaderLocation.TOP:
                    raw_headers = [
                        cell.value for cell in raw_ws[table.range.start_cell.row]
                    ]
                    reduced_headers = [
                        cell.value for cell in reduced_ws[table.range.start_cell.row]
                    ]
                else:
                    raw_headers = [
                        raw_ws[cell.coordinate].value
                        for cell in raw_ws[table.range.start_cell.column]
                    ]
                    reduced_headers = [
                        reduced_ws[cell.coordinate].value
                        for cell in reduced_ws[table.range.start_cell.column]
                    ]

                if raw_headers != reduced_headers:
                    print(raw_headers)
                    print(reduced_headers)
                    return False

        return True
