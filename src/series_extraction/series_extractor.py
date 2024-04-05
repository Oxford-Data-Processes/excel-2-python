from openpyxl.utils import get_column_letter
from objects import (
    HeaderLocation,
    ExcelFile,
    Worksheet,
    Table,
    Series,
    Cell,
    SeriesDataType,
)

from typing import Dict, List

from uuid import UUID, uuid4


class SeriesExtractor:

    @staticmethod
    def build_series_data_top(
        data, sheet, header, header_location, start_row, end_row, col_index
    ):
        series_data = {
            "series_header": header,
            "row_formulas": [],
            "row_values": [],
            "header_location": header_location,
            "series_starting_cell": {"row": start_row + 1, "column": col_index},
            "series_length": f"{end_row - start_row}",
        }
        for row in range(start_row + 1, start_row + 3):
            cell = data[sheet.sheet_name][f"{get_column_letter(col_index)}{row}"]
            series_data["row_formulas"].append(cell.get("formula"))
            series_data["row_values"].append(cell.get("value"))

        if series_data["row_values"]:
            series_data["data_type"] = type(series_data["row_values"][0]).__name__

        return series_data

    @staticmethod
    def build_series_data_left(
        data, sheet, header, header_location, start_column, end_column, row_index
    ):
        series_data = {
            "series_header": header,
            "row_formulas": [],
            "row_values": [],
            "header_location": header_location,
            "series_starting_cell": {"row": row_index, "column": start_column + 1},
            "series_length": f"{end_column - start_column}",
        }
        for col_offset in range(1, 3):
            cell = data[sheet.sheet_name][
                f"{get_column_letter(start_column + col_offset)}{row_index}"
            ]
            series_data["row_formulas"].append(cell.get("formula"))
            series_data["row_values"].append(cell.get("value"))

        if series_data["row_values"]:
            series_data["data_type"] = type(series_data["row_values"][0]).__name__

        return series_data

    @staticmethod
    def extract_table_details(located_tables, data):
        tables_data = {}

        for sheet, tables in located_tables.items():
            sheet_data = {}
            for table in tables:
                start_column = table.range.start_cell.column
                end_column = table.range.end_cell.column
                start_row = table.range.start_cell.row
                end_row = table.range.end_cell.row
                header_location = table.header_location
                header_values = table.header_values

                if header_location == HeaderLocation.TOP:
                    for col_index, header in enumerate(
                        header_values, start=start_column
                    ):
                        range_identifier = f"{get_column_letter(col_index)}{start_row + 1}:{get_column_letter(col_index)}{end_row}"
                        series_data = SeriesExtractor.build_series_data_top(
                            data,
                            sheet,
                            header,
                            header_location.value,
                            start_row,
                            end_row,
                            col_index,
                        )
                        sheet_data[range_identifier] = series_data
                else:
                    for row_index, header in enumerate(header_values, start=start_row):
                        row_range_identifier = f"{get_column_letter(start_column + 1)}{row_index}:{get_column_letter(end_column)}{row_index}"

                        series_data = SeriesExtractor.build_series_data_left(
                            data,
                            sheet,
                            header,
                            header_location.value,
                            start_column,
                            end_column,
                            row_index,
                        )
                        sheet_data[row_range_identifier] = series_data

            tables_data[sheet] = sheet_data

        return tables_data

    @staticmethod
    def extract_series(
        extracted_tables: Dict[Worksheet, List[Table]],
        data: dict,
    ) -> Dict[UUID, List[Series]]:
        series_data = SeriesExtractor.extract_table_details(extracted_tables, data)

        series = {}

        for worksheet, table_data in series_data.items():
            series[worksheet.sheet_name] = []
            for _, series_data in table_data.items():
                series[worksheet.sheet_name].append(
                    Series(
                        series_id=uuid4(),
                        worksheet=worksheet,
                        series_header=series_data["series_header"],
                        formulas=series_data["row_formulas"],
                        values=series_data["row_values"],
                        header_location=HeaderLocation(series_data["header_location"]),
                        series_starting_cell=Cell(
                            column=series_data["series_starting_cell"]["column"],
                            row=series_data["series_starting_cell"]["row"],
                            coordinate=f"{get_column_letter(series_data['series_starting_cell']['column'])}{series_data['series_starting_cell']['row']}",
                        ),
                        series_length=int(series_data["series_length"]),
                        series_data_type=SeriesDataType(series_data["data_type"]),
                    )
                )

        return series
