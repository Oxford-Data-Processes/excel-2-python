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
        orientation: str,
    ):
        series = {
            "series_header": header,
            "formulas": [],
            "values": [],
            "header_location": header_location,
            "series_starting_cell": (
                {"row": start_row_or_column + 1, "column": index}
                if orientation == "top"
                else {"row": index, "column": start_row_or_column + 1}
            ),
            "series_length": f"{end_row_or_column - start_row_or_column}",
        }
        sheet_data = workbook_data.get_sheet_data(sheet.sheet_name)
        for offset in range(1, 3):
            cell_key = (
                f"{get_column_letter(index)}{start_row_or_column + offset}"
                if orientation == "top"
                else f"{get_column_letter(start_row_or_column + offset)}{index}"
            )
            cell = sheet_data[cell_key]
            series["formulas"].append(cell.formula)
            series["values"].append(cell.value)

        if series["values"]:
            series["data_type"] = type(series["values"][0]).__name__

        return series

    @staticmethod
    def handle_top(
        start_column,
        end_column,
        start_row,
        end_row,
        header_location,
        header_values,
        sheet_data,
        workbook_data,
        sheet,
    ):
        for col_index, header in enumerate(header_values, start=start_column):
            range_identifier = f"{get_column_letter(col_index)}{start_row + 1}:{get_column_letter(col_index)}{end_row}"
            series = SeriesExtractor.build_series(
                workbook_data=workbook_data,
                sheet=sheet,
                header=header,
                header_location=header_location,
                start_row_or_column=start_row,
                end_row_or_column=end_row,
                index=col_index,
                orientation="top",
            )
            sheet_data[range_identifier] = series

    @staticmethod
    def handle_left(
        start_column,
        end_column,
        start_row,
        end_row,
        header_location,
        header_values,
        sheet_data,
        workbook_data,
        sheet,
    ):
        for row_index, header in enumerate(header_values, start=start_row):
            row_range_identifier = f"{get_column_letter(start_column + 1)}{row_index}:{get_column_letter(end_column)}{row_index}"
            series = SeriesExtractor.build_series(
                workbook_data=workbook_data,
                sheet=sheet,
                header=header,
                header_location=header_location,
                start_row_or_column=start_column,
                end_row_or_column=end_column,
                index=row_index,
                orientation="left",
            )
            sheet_data[row_range_identifier] = series

    @staticmethod
    def extract_table_details(extracted_tables, workbook_data):
        tables_data = {}

        for sheet, tables in extracted_tables.items():
            sheet_data = {}
            for table in tables:
                start_column = table.range.start_cell.column
                end_column = table.range.end_cell.column
                start_row = table.range.start_cell.row
                end_row = table.range.end_cell.row
                header_location = table.header_location
                header_values = table.header_values

                if header_location == HeaderLocation.TOP:
                    SeriesExtractor.handle_top(
                        start_column,
                        end_column,
                        start_row,
                        end_row,
                        header_location,
                        header_values,
                        sheet_data,
                        workbook_data,
                        sheet,
                    )
                else:
                    SeriesExtractor.handle_left(
                        start_column,
                        end_column,
                        start_row,
                        end_row,
                        header_location,
                        header_values,
                        sheet_data,
                        workbook_data,
                        sheet,
                    )

                tables_data[sheet] = sheet_data

        return tables_data

    @staticmethod
    def calculate_header_cell(series):
        series_starting_cell_column = series["series_starting_cell"]["column"]
        series_starting_cell_row = series["series_starting_cell"]["row"]

        if series["header_location"] == HeaderLocation.TOP:
            header_cell_row = series_starting_cell_row - 1
            header_cell_column = series_starting_cell_column
        elif series["header_location"] == HeaderLocation.LEFT:
            header_cell_row = series_starting_cell_row
            header_cell_column = series_starting_cell_column - 1

        return header_cell_row, header_cell_column

    @staticmethod
    def extract_series(
        extracted_tables: Dict[Worksheet, List[Table]],
        workbook_data: dict,
    ) -> Dict[str, List[Series]]:
        series = SeriesExtractor.extract_table_details(
            extracted_tables, workbook_data
        )

        series_collection = {}

        for worksheet, table_data in series.items():
            series_collection[worksheet.sheet_name] = []
            for _, series in table_data.items():
                header_cell_row, header_cell_column = (
                    SeriesExtractor.calculate_header_cell(series)
                )
                series_obj = Series(
                    series_id=SeriesId(
                        sheet_name=worksheet.sheet_name,
                        series_header=series["series_header"],
                        series_header_cell_row=header_cell_row,
                        series_header_cell_column=header_cell_column,
                    ),
                    worksheet=worksheet,
                    series_header=series["series_header"],
                    formulas=series["formulas"],
                    values=series["values"],
                    header_location=series["header_location"],
                    series_starting_cell=Cell(
                        column=series["series_starting_cell"]["column"],
                        row=series["series_starting_cell"]["row"],
                        coordinate=f"{get_column_letter(series['series_starting_cell']['column'])}{series['series_starting_cell']['row']}",
                    ),
                    series_length=int(series["series_length"]),
                    series_data_type=SeriesDataType(series["data_type"]),
                )
                series_collection[worksheet.sheet_name].append(series_obj)

        return series_collection
