from __future__ import annotations

from abc import ABC, abstractmethod
from functools import lru_cache
from pathlib import Path
from typing import Dict, List

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
    @lru_cache
    def dictionary(self) -> pd.DataFrame:
        """All cell couples in document"""
        data = []
        for table in self.tables:
            data.extend(self.table_to_dict(table))

        return pd.DataFrame(data)

    @staticmethod
    def table_to_dict(table: pd.DataFrame) -> List[Dict[str, str]]:
        pairs = []
        cols = table.columns.to_list()
        data = table.to_numpy()

        for k, col in enumerate(cols[:-1]):
            pair = {'title': col, 'content': cols[k+1], 'orientation': 'row'}
            pairs.append(pair)

        for col, value in zip(cols, data[0]):
            pair = {'title': col, 'content': value, 'orientation': 'column'}
            pairs.append(pair)

        for row in data:
            for k, value in enumerate(row[:-1]):
                pair = {'title': value, 'content': value[k+1], 'orientation': 'row'}
                pairs.append(pair)

        for row in data.T:
            for k, value in enumerate(row[:-1]):
                pair = {'title': value, 'content': value[k+1], 'orientation': 'column'}
                pairs.append(pair)


        return pairs


class BaseNode(BaseHandler):
    """The Base Node is an abstract class that defines the interface for all other nodes.
    This interface will be used by the Extractor to search for tables in the document.
    """

    nodes: List[BaseNode] = []

    def __init__(self, title: str) -> None:
        self.title = title
        self.nodes = []

    def get_words(self, recursive: bool) -> List[str]:
        if not recursive:
            return self.words
        else:
            words = self.words
            for node in self.nodes:
                words += node.get_words(recursive=True)
            return words

    def get_tables(self, recursive: bool) -> List[pd.DataFrame]:
        if not recursive:
            return self.tables
        else:
            tables = self.tables
            for node in self.nodes:
                tables += node.get_tables(recursive=True)
            return tables

    def get_dictionary(self, recursive: bool) -> pd.DataFrame:
        tables = self.get_tables(recursive=recursive)

        data = {}
        for table in tables:



class TreeFileHandler(BaseHandler):
    @property
    def root(self) -> BaseNode:
        return BaseNode("")
