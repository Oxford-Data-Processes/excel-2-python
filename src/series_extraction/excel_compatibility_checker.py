from objects import ExcelFile, Worksheet, Table
from typing import Dict


class ExcelCompatibilityChecker:
    @staticmethod
    def check_file(
        excel_raw: ExcelFile,
        excel_reduced: ExcelFile,
        extracted_tables: Dict[Worksheet, Table],
    ) -> bool:

        is_compatible = True
        return is_compatible
