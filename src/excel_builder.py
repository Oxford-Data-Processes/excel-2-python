import openpyxl


class ExcelBuilder:
    def __init__(self, series_list, output_file_path: str):
        self.series_list = series_list
        self.output_file_path = output_file_path

    def create_excel_from_series(self):
        # Initialize a new workbook and dictionary to track existing sheets
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Start with a clean slate by removing the default sheet
        ws_dict = {}

        # Iterate through each series object
        for series in self.series_list:
            sheet_name = series.series_id.sheet_name
            series_header = series.series_header
            header_row = series.series_id.series_header_cell_row
            header_col = series.series_id.series_header_cell_column
            header_location = series.header_location
            formulas = series.formulas
            values = series.values
            start_row = series.series_starting_cell.row
            start_col = series.series_starting_cell.column

            # Ensure the worksheet exists or create it if not
            if sheet_name not in ws_dict:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                else:
                    ws = wb.create_sheet(title=sheet_name)
                ws_dict[sheet_name] = ws
            ws = ws_dict[sheet_name]

            # Place the series header
            ws.cell(row=header_row, column=header_col, value=series_header)

            # Fill the cells with formulas or values
            for i in range(series.series_length):

                row = start_row + i if header_location.value == "top" else start_row
                col = start_col + i if header_location.value == "left" else start_col

                if formulas != [None, None]:
                    cell = ws.cell(row=row, column=col)
                    cell.value = formulas[i]
                else:
                    cell = ws.cell(row=row, column=col)
                    cell.value = values[i]

        # Save the workbook to a file
        wb.save(self.output_file_path)
