import openpyxl


class ExcelBuilder:

    @staticmethod
    def create_excel_from_workbook(workbook: openpyxl.Workbook, output_file_path: str):
        # Create a new workbook
        new_workbook = openpyxl.Workbook()
        new_workbook.remove(
            new_workbook.active
        )  # Start with a clean slate by removing the default sheet

        # Iterate over all sheets in the original workbook
        for sheet in workbook:
            # Create a new sheet in the new workbook with the same name
            new_sheet = new_workbook.create_sheet(title=sheet.title)

            # Iterate over all cells in the original sheet
            for row in sheet.iter_rows():
                for cell in row:
                    # Copy all cell values to the new sheet
                    new_cell = new_sheet.cell(row=cell.row, column=cell.column)
                    new_cell.value = cell.value

        # Save the new workbook to a file
        new_workbook.save(output_file_path)

    @staticmethod
    def create_excel_from_series(series_list, output_file_path, values_only=False):
        # Initialize a new workbook and dictionary to track existing sheets
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Start with a clean slate by removing the default sheet
        ws_dict = {}

        # Iterate through each series object
        for series in series_list:
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

                if values_only is False:  # Fill with formulas
                    if formulas != [None, None]:
                        cell = ws.cell(row=row, column=col)
                        cell.value = formulas[i]
                    else:
                        cell = ws.cell(row=row, column=col)
                        cell.value = values[i]
                else:
                    cell = ws.cell(row=row, column=col)
                    cell.value = values[i]

        # Save the workbook to a file
        wb.save(output_file_path)
