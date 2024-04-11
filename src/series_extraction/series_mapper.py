from typing import List, Tuple

from objects import Cell, Series, Worksheet


class SeriesMapper:

    @staticmethod
    def get_cells(series: Series) -> List[Cell]:
        """Get list of all Excel cells in a series"""
        cells = []
        offset = range(series.series_length)
        is_top_header = series.header_location.value == "top"
        is_left_header = series.header_location.value == "left"

        for i in offset:
            cell = Cell(
                column=series.series_starting_cell.column
                + (i if is_left_header else 0),
                row=series.series_starting_cell.row + (i if is_top_header else 0),
                coordinate=None,
                value=None,
                value_type=None,
            )
            cells.append(cell)
        return cells

    @staticmethod
    def map_series(
        series_dict: dict[str, Series]
    ) -> dict[Worksheet, dict[Cell, Tuple[int, Series]]]:
        """Get mapping from Cell to row index and Series"""
        series_mapping = {}
        for sheet_name, series in series_dict.items():
            worksheet = Worksheet(
                sheet_name=sheet_name, workbook_file_path=None, worksheet=None
            )
            series_mapping[worksheet] = {}
            for _, s in enumerate(series):
                cells = SeriesMapper.get_cells(s)
                for index, cell in enumerate(cells):
                    series_mapping[worksheet][cell] = (index, s)
        return series_mapping
