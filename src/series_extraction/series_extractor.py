from openpyxl.utils import get_column_letter
from objects import (
    HeaderLocation,
    Worksheet,
    Table,
    Cell,
    Series,
    SeriesId,
    SeriesDataType,
)

from typing import Dict, List


# class SeriesDataType(Enum):
#     INT = "int"
#     STR = "str"
#     FLOAT = "float"
#     BOOL = "bool"
#     TIME = "time"


# @dataclass(frozen=True)
# class SeriesId:
#     sheet_name: str
#     series_header: str
#     series_header_cell_row: int
#     series_header_cell_column: int

#     def __str__(self):
#         return "|".join(
#             [
#                 self.sheet_name,
#                 self.series_header,
#                 str(self.series_header_cell_row),
#                 str(self.series_header_cell_column),
#             ]
#         )


# @dataclass
# class Series:
#     series_id: SeriesId
#     worksheet: Worksheet
#     series_header: str
#     formulas: List[str]
#     values: List[Union[int, str, float, bool]]
#     header_location: HeaderLocation
#     series_starting_cell: Cell
#     series_length: int
#     series_data_type: SeriesDataType


class SeriesExtractor:

    @staticmethod
    def build_series(
        workbook_data,
        sheet,
        header,
        header_location,
        start_row_or_column,
        end_row_or_column,
        index,
        orientation,
    ):
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

        series_id = SeriesId(
            sheet_name=sheet.sheet_name,
            series_header=header,
            series_header_cell_row=(
                start_row_or_column if orientation == "top" else index
            ),
            series_header_cell_column=(
                index if orientation == "top" else start_row_or_column + 1
            ),
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
                coordinate=f"{get_column_letter(start_cell_column)}{start_cell_row}",
            ),
            series_length=end_row_or_column - start_row_or_column,
            series_data_type=series_data_type,
        )

    @staticmethod
    def handle_series(
        workbook_data,
        sheet,
        header_values,
        start_column,
        end_column,
        start_row,
        end_row,
        header_location,
    ):
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
    def extract_table_details(extracted_tables, workbook_data):
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
    def calculate_header_cell(column, row, header_location):
        return (
            (row - 1, column)
            if header_location == HeaderLocation.TOP
            else (row, column - 1)
        )

    @staticmethod
    def extract_series(
        extracted_tables,
        workbook_data,
    ):
        detailed_series = SeriesExtractor.extract_table_details(
            extracted_tables, workbook_data
        )

        series_collection = {}

        for worksheet_obj, table_data in detailed_series.items():
            series_collection[worksheet_obj.sheet_name] = []
            for _, single_series in table_data.items():
                column, row = (
                    single_series.series_starting_cell.column,
                    single_series.series_starting_cell.row,
                )
                header_cell_row, header_cell_column = (
                    SeriesExtractor.calculate_header_cell(
                        column, row, single_series.header_location
                    )
                )
                series_instance = Series(
                    series_id=SeriesId(
                        sheet_name=single_series.worksheet.sheet_name,
                        series_header=single_series.series_header,
                        series_header_cell_row=header_cell_row,
                        series_header_cell_column=header_cell_column,
                    ),
                    worksheet=worksheet_obj,
                    series_header=single_series.series_header,
                    formulas=single_series.formulas,
                    values=single_series.values,
                    header_location=single_series.header_location,
                    series_starting_cell=Cell(
                        column=single_series.series_starting_cell.column,
                        row=single_series.series_starting_cell.row,
                        coordinate=f"{get_column_letter(single_series.series_starting_cell.column)}{single_series.series_starting_cell.row}",
                    ),
                    series_length=int(single_series.series_length),
                    series_data_type=SeriesDataType(single_series.series_data_type),
                )
                series_collection[worksheet_obj.sheet_name].append(series_instance)

        return series_collection
