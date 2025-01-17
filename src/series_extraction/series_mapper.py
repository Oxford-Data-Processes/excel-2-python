from typing import List, Tuple

from objects import Cell, Series, SeriesId, Worksheet


class SeriesMapper:

    @staticmethod
    def get_cells(series: Series) -> List[Cell]:
        """Get list of all Excel cells in a series"""
        cells = []
        offset = range(series.series_length)

        for i in offset:
            cell = Cell(
                column=series.series_starting_cell.column + i,
                row=series.series_starting_cell.row + i,
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
            worksheet = Worksheet(sheet_name=sheet_name)
            series_mapping[worksheet] = {}
            for _, s in enumerate(series):
                cells = SeriesMapper.get_cells(s)
                for index, cell in enumerate(cells):
                    series_mapping[worksheet][cell] = (index, s)
        return series_mapping

    @staticmethod
    def get_series_from_series_id(
        series_id: SeriesId, series_dict: dict[str, List[Series]]
    ) -> Series:
        sheet_name = series_id.sheet_name
        series_list = series_dict[sheet_name]
        series = next(series for series in series_list if series.series_id == series_id)
        return series
