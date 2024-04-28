from objects import ExcelFile, Worksheet, Table, Cell, CellRange, HeaderLocation
from excel_utils import ExcelUtils
from typing import Dict, List, Set, Tuple


class TableFinder:
    @staticmethod
    def _cells_are_adjacent(cell1: Cell, cell2: Cell) -> bool:
        """Check if two cells are adjacent based on their row and column indices."""
        return abs(cell1.row - cell2.row) <= 1 and abs(cell1.column - cell2.column) <= 1

    @staticmethod
    def _extract_non_empty_cells(sheet_data: Dict) -> Set[Cell]:
        """Extract non-empty cells from the sheet data."""
        return {
            Cell(row=data["row"], column=data["column"], value=data["value"])
            for data in sheet_data.values()
            if data["value"] is not None
        }

    @staticmethod
    def _process_frontier(frontier: Set[Cell], non_empty_cells: Set[Cell]) -> Set[Cell]:
        """Identify and process all adjacent cells for the given frontier."""
        new_frontier = set()
        for frontier_cell in frontier:
            adjacent_cells = {
                cell
                for cell in non_empty_cells
                if TableFinder._cells_are_adjacent(cell, frontier_cell)
            }
            non_empty_cells.difference_update(adjacent_cells)
            new_frontier.update(adjacent_cells)
        return new_frontier

    @staticmethod
    def _update_cluster(
        cluster: Set[Cell], new_frontier: Set[Cell]
    ) -> Tuple[Set[Cell], int, int, int, int]:
        """Update the cluster with new cells and adjust the boundaries."""
        cluster.update(new_frontier)
        min_row = min(cluster, key=lambda x: x.row).row
        max_row = max(cluster, key=lambda x: x.row).row
        min_col = min(cluster, key=lambda x: x.column).column
        max_col = max(cluster, key=lambda x: x.column).column
        return cluster, min_row, max_row, min_col, max_col

    @staticmethod
    def _expand_cluster(
        frontier: Set[Cell], non_empty_cells: Set[Cell]
    ) -> Tuple[Set[Cell], int, int, int, int]:
        """Expand the cluster from the frontier cells, updating boundaries."""
        cluster = set(frontier)
        while frontier:
            new_frontier = TableFinder._process_frontier(frontier, non_empty_cells)
            cluster, min_row, max_row, min_col, max_col = TableFinder._update_cluster(
                cluster, new_frontier
            )
            frontier = new_frontier

        return cluster, min_row, max_row, min_col, max_col

    @staticmethod
    def find_table_boundaries(sheet_data: Dict) -> List[Tuple[int, int, int, int]]:
        """Identify table boundaries by clustering adjacent non-empty cells."""
        non_empty_cells = TableFinder._extract_non_empty_cells(sheet_data)
        tables = []

        while non_empty_cells:
            initial_cell = non_empty_cells.pop()
            cluster, min_row, max_row, min_col, max_col = TableFinder._expand_cluster(
                {initial_cell}, non_empty_cells
            )
            tables.append((min_row, min_col, max_row, max_col))

        return tables

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
                            "coordinate": f"{ExcelUtils.get_column_letter_from_number(bound[1])}{bound[0]}",
                        },
                        "end_cell": {
                            "column": bound[3],
                            "row": bound[2],
                            "coordinate": f"{ExcelUtils.get_column_letter_from_number(bound[3])}{bound[2]}",
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
            column_letter = ExcelUtils.get_column_letter_from_number(column)
            cell_coordinate = f"{column_letter}{start_row}"
            if data_object[cell_coordinate]["value_type"] != "str":
                return False, None
            else:
                header_values.append(data_object[cell_coordinate]["value"])

        return True, header_values

    @staticmethod
    def get_first_column_values(range_input, data_object):
        start_row = range_input["range"]["start_cell"]["row"]
        end_row = range_input["range"]["end_cell"]["row"]
        start_column = range_input["range"]["start_cell"]["column"]

        first_column_values = []

        for row in range(start_row, end_row + 1):
            column_letter = ExcelUtils.get_column_letter_from_number(start_column)
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
