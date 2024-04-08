from typing import Dict, List, Iterator, Any
from uuid import UUID

from objects import Series


class SeriesIterator:
    @staticmethod
    def iterate_series(series: Dict[UUID, List[Series]]) -> Iterator[Series]:
        for series_list in series.values():
            for single_series in series_list:
                yield single_series
