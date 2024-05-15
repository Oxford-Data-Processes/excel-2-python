import openpyxl

from typing import List
from objects import Series


class ExcelDataExtractor:
    @staticmethod
    def extract_series_data_from_excel(
        workbook: openpyxl.workbook.workbook, series_list: List[Series]
    ):
        series_values_dict_raw = {}

        for series in series_list:
            series_id = series.series_id
            header_location = series.header_location

            sheet_name = series_id.sheet_name
            series_header_cell_row = series_id.series_header_cell_row
            series_header_cell_column = series_id.series_header_cell_column

            # Access the sheet
            sheet = workbook[sheet_name]

            # Extract data based on header location
            if header_location.value.lower() == "top":
                # Get values below the series header cell
                data = []
                row = series_header_cell_row + 1
                while row <= sheet.max_row:
                    cell_value = sheet.cell(
                        row=row, column=series_header_cell_column
                    ).value
                    data.append(cell_value)
                    row += 1
            elif header_location.value.lower() == "left":
                data = []
                column = series_header_cell_column + 1
                while column <= sheet.max_column:
                    cell_value = sheet.cell(
                        row=series_header_cell_row, column=column
                    ).value
                    data.append(cell_value)
                    column += 1
            else:
                raise ValueError(f"Unknown header location: {header_location.value}")

            series_values_dict_raw[str(series_id)] = data

        return series_values_dict_raw
