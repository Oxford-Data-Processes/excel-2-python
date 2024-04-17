from objects import ExcelFile, Worksheet, Table, Cell, CellRange, HeaderLocation

from typing import Dict, List


class TableFinder:

    @staticmethod
    def find_table_boundaries(sheet_data: dict):
        non_empty_cells = {
            (cell_data["row"], cell_data["column"])
            for cell_data in sheet_data.values()
            if cell_data["value"] is not None
        }

        def are_adjacent(cell1, cell2):
            return abs(cell1[0] - cell2[0]) <= 1 and abs(cell1[1] - cell2[1]) <= 1

        tables = []
        while non_empty_cells:
            # Initialize with a single cell
            cell = non_empty_cells.pop()
            cluster = {cell}
            frontier = {cell}

            # Initialize boundaries
            min_row, max_row = cell[0], cell[0]
            min_col, max_col = cell[1], cell[1]

            while frontier:
                new_frontier = set()
                for frontier_cell in frontier:
                    # Check adjacent cells
                    for cell in set(non_empty_cells):
                        if are_adjacent(cell, frontier_cell):
                            # Update boundaries
                            min_row, max_row = min(min_row, cell[0]), max(
                                max_row, cell[0]
                            )
                            min_col, max_col = min(min_col, cell[1]), max(
                                max_col, cell[1]
                            )

                            cluster.add(cell)
                            new_frontier.add(cell)
                            non_empty_cells.remove(cell)

                frontier = new_frontier

            tables.append((min_row, min_col, max_row, max_col))

        return tables

    @staticmethod
    def column_number_to_letter(column_number):
        letter = ""
        while column_number > 0:
            column_number, remainder = divmod(column_number - 1, 26)
            letter = chr(65 + remainder) + letter
        return letter

    @staticmethod
    def locate_data_tables(data) -> dict:
        located_tables = {}

        for sheet_name, sheet_data in data.items():
            table_boundaries = TableFinder.find_table_boundaries(sheet_data)
            located_tables[sheet_name] = [
                {
                    "name": f"{sheet_name}_{index+1}",
                    "range": {
                        "start_cell": {
                            "column": bound[1],
                            "row": bound[0],
                            "coordinate": f"{TableFinder.column_number_to_letter(bound[1])}{bound[0]}",
                        },
                        "end_cell": {
                            "column": bound[3],
                            "row": bound[2],
                            "coordinate": f"{TableFinder.column_number_to_letter(bound[3])}{bound[2]}",
                        },
                    },
                }
                for index, bound in enumerate(table_boundaries)
            ]

        return located_tables

    @staticmethod
    def are_first_row_values_strings(range_input, data_object):

        start_column = range_input["range"]["start_cell"]["column"]
        start_row = range_input["range"]["start_cell"]["row"]
        end_column = range_input["range"]["end_cell"]["column"]

        header_values = []

        for column in range(start_column, end_column + 1):
            column_letter = TableFinder.column_number_to_letter(column)
            cell_coordinate = f"{column_letter}{start_row}"
            if data_object[cell_coordinate]["value_type"] != "str":
                return False, None
            else:
                header_values.append(data_object[cell_coordinate]["value"])

        print("header values")
        print(header_values)

        return True, header_values

    @staticmethod
    def get_first_column_values(range_input, data_object):
        start_row = range_input["range"]["start_cell"]["row"]
        end_row = range_input["range"]["end_cell"]["row"]
        start_column = range_input["range"]["start_cell"]["column"]

        first_column_values = []

        for row in range(start_row, end_row + 1):
            column_letter = TableFinder.column_number_to_letter(start_column)
            cell_coordinate = f"{column_letter}{row}"
            first_column_values.append(data_object[cell_coordinate]["value"])

        return first_column_values

    @staticmethod
    def get_header_location_and_values(data):

        located_tables = TableFinder.locate_data_tables(data)

        for sheet, table_details in located_tables.items():
            for item in table_details:
                boolean, header_values = TableFinder.are_first_row_values_strings(
                    item, data[sheet]
                )
                if boolean:
                    item["header_location"] = "top"
                    item["header_values"] = header_values
                else:
                    item["header_location"] = "left"
                    item["header_values"] = TableFinder.get_first_column_values(
                        item, data[sheet]
                    )

        return located_tables

    @staticmethod
    def find_tables(
        excel_file: ExcelFile,
    ) -> Dict[Worksheet, List[Table]]:
        data = {}
        for ws_values, ws_formulas in zip(
            excel_file.workbook_with_values.worksheets,
            excel_file.workbook_with_formulas.worksheets,
        ):
            worksheet = Worksheet(
                sheet_name=ws_values.title,
                workbook_file_path=None,
                worksheet=ws_values,
            )
            sheet_data = {}
            for row_values, row_formulas in zip(
                ws_values.iter_rows(), ws_formulas.iter_rows()
            ):
                for cell_values, cell_formulas in zip(row_values, row_formulas):
                    cell_coordinate = cell_values.coordinate
                    cell_value = cell_values.value
                    cell_value_type = type(cell_value).__name__
                    formula = (
                        cell_formulas.value
                        if isinstance(cell_formulas.value, str)
                        and cell_formulas.value.startswith("=")
                        else None
                    )
                    sheet_data[cell_coordinate] = {
                        "value": cell_value,
                        "value_type": cell_value_type,
                        "row": cell_values.row,
                        "column": cell_values.column,
                        "coordinate": cell_coordinate,
                        "formula": formula,
                    }
            data[worksheet.sheet_name] = sheet_data

        located_tables = TableFinder.get_header_location_and_values(data)

        extracted_tables = {}
        for sheet_name, tables in located_tables.items():
            worksheet_tables = []
            for table in tables:
                range_start = Cell(
                    column=table["range"]["start_cell"]["column"],
                    row=table["range"]["start_cell"]["row"],
                    coordinate=table["range"]["start_cell"]["coordinate"],
                )
                range_end = Cell(
                    column=table["range"]["end_cell"]["column"],
                    row=table["range"]["end_cell"]["row"],
                    coordinate=table["range"]["end_cell"]["coordinate"],
                )
                cell_range = CellRange(start_cell=range_start, end_cell=range_end)
                header_location = HeaderLocation(table["header_location"])
                table_obj = Table(
                    name=table["name"],
                    range=cell_range,
                    header_location=header_location,
                    header_values=table["header_values"],
                )
                worksheet_tables.append(table_obj)
            extracted_tables[
                Worksheet(
                    sheet_name=sheet_name,
                    workbook_file_path=None,
                    worksheet=None,
                )
            ] = worksheet_tables

        return extracted_tables, data
