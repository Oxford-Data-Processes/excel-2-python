from objects import (
    HeaderLocation,
    Worksheet,
    WorkbookData,
    Table,
    Cell,
    Series,
    SeriesId,
    SeriesDataType,
)
from openpyxl.utils import get_column_letter

from typing import Dict, List


class SeriesExtractor:

    @staticmethod
    def build_series(
        workbook_data: WorkbookData,
        sheet: Worksheet,
        header: str,
        header_location: HeaderLocation,
        start_row_or_column: int,
        end_row_or_column: int,
        index: int,
        orientation: str,
    ) -> Series:
        start_cell_row = start_row_or_column + 1 if orientation == "top" else index
        start_cell_column = index if orientation == "top" else start_row_or_column + 1

        row_formulas, row_values = [], []
        sheet_data = workbook_data.get_sheet_data(sheet.sheet_name)
        for offset in range(1, 3):
            cell_key = (
                f"{get_column_letter(index)}{start_row_or_column + offset}"
                if orientation == "top"
                else f"{get_column_letter(start_row_or_column + offset)}{index}"
            )
            cell = sheet_data.get(cell_key)
            row_formulas.append(cell.formula if cell else None)
            row_values.append(cell.value if cell else None)

        series_header_cell_row, series_header_cell_column = (
            SeriesExtractor.calculate_header_cell(
                start_cell_row, start_cell_column, header_location
            )
        )

        series_id = SeriesId(
            sheet_name=sheet.sheet_name,
            series_header=header,
            series_header_cell_row=series_header_cell_row,
            series_header_cell_column=series_header_cell_column,
        )

        series_data_type = (
            SeriesDataType(type(row_values[0]).__name__)
            if row_values
            else SeriesDataType.STR
        )

        return Series(
            series_id=series_id,
            worksheet=sheet,
            series_header=header,
            formulas=row_formulas,
            values=row_values,
            header_location=header_location,
            series_starting_cell=Cell(
                row=start_cell_row,
                column=start_cell_column,
            ),
            series_length=end_row_or_column - start_row_or_column,
            series_data_type=series_data_type,
        )

    @staticmethod
    def handle_series(
        workbook_data: WorkbookData,
        sheet: Worksheet,
        header_values: List[str],
        start_column: int,
        end_column: int,
        start_row: int,
        end_row: int,
        header_location: HeaderLocation,
    ) -> Dict[str, Series]:
        orientation = "top" if header_location == HeaderLocation.TOP else "left"
        series_data = {}
        for index, header in enumerate(
            header_values, start=start_column if orientation == "top" else start_row
        ):
            range_identifier = (
                f"{get_column_letter(index)}{start_row}:{get_column_letter(index)}{end_row}"
                if orientation == "top"
                else f"{get_column_letter(start_column)}{index}:{get_column_letter(end_column)}{index}"
            )
            series = SeriesExtractor.build_series(
                workbook_data,
                sheet,
                header,
                header_location,
                start_row if orientation == "top" else start_column,
                end_row if orientation == "top" else end_column,
                index,
                orientation,
            )
            series_data[range_identifier] = series
        return series_data

    @staticmethod
    def extract_table_details(
        extracted_tables: Dict[Worksheet, List[Table]], workbook_data: WorkbookData
    ):
        tables_data = {}
        for sheet, tables in extracted_tables.items():
            sheet_data = {}
            for table in tables:
                series_result = SeriesExtractor.handle_series(
                    workbook_data,
                    sheet,
                    table.header_values,
                    table.range.start_cell.column,
                    table.range.end_cell.column,
                    table.range.start_cell.row,
                    table.range.end_cell.row,
                    table.header_location,
                )
                sheet_data.update(series_result)
            tables_data[sheet] = sheet_data
        return tables_data

    @staticmethod
    def calculate_header_cell(
        series_starting_cell_row: int,
        series_starting_cell_column: int,
        header_location: HeaderLocation,
    ):
        return (
            (series_starting_cell_row - 1, series_starting_cell_column)
            if header_location == HeaderLocation.TOP
            else (series_starting_cell_row, series_starting_cell_column - 1)
        )

    @staticmethod
    def extract_series(
        extracted_tables: Dict[Worksheet, List[Table]],
        workbook_data: WorkbookData,
    ) -> Dict[str, List[Series]]:
        detailed_series = SeriesExtractor.extract_table_details(
            extracted_tables, workbook_data
        )

        series_collection = {}

        for worksheet, table_data in detailed_series.items():
            series_collection[worksheet.sheet_name] = []
            for _, single_series in table_data.items():
                series_collection[worksheet.sheet_name].append(single_series)

        return series_collection
