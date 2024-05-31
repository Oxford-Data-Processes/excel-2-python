from typing import List, Dict, Tuple

from objects import Series, SeriesId


class SeriesMapper:

    @staticmethod
    def map_series(
        series_dict: Dict[str, List[Series]]
    ) -> Dict[Tuple[str, int], Series]:
        """Map from (sheet_name, column_number) to Series"""
        series_mapping = {}
        for sheet_name, series_list in series_dict.items():
            for series in series_list:
                column_number = series.series_starting_cell.column
                series_mapping[(sheet_name, column_number)] = series
        return series_mapping

    @staticmethod
    def get_series_from_series_id(
        series_id: SeriesId, series_dict: Dict[str, List[Series]]
    ) -> Series:
        """Retrieve a series based on its unique identifier within the dictionary"""
        return next(
            (
                series
                for series in series_dict[series_id.sheet_name]
                if series.series_id == series_id
            ),
            None,  # Returns None if no series matches
        )
