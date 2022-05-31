from abc import ABC
from pathlib import Path
from typing import List

import pandas as pd


class BaseHandler(ABC):
    """The Base Handler is an abstract class that defines the interface for all other handlers.
    This interface will be used by the Extractor to search for tables in the document.
    """

    def __init__(self, file_path: Path) -> None:
        self.file_path = file_path

    @property
    def words(self) -> List[str]:
        """List of all words in document"""
        return []

    @property
    def tables(self) -> List[pd.DataFrame]:
        """List of all tables (as dataframes) in document"""
        return []

    @property
    def dictionary(self) -> pd.DataFrame:
        """All cell couples in document"""
        return pd.DataFrame()
