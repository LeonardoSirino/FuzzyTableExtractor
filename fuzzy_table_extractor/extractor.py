from dataclasses import dataclass
from enum import Enum, auto
from typing import Callable, List

import numpy as np
import pandas as pd

from .doc_handlers import BaseHandler
from .util import match_regex_list, str_comparison


@dataclass
class TableMatch:
    search_term: str
    original_term: str
    score: str


class FieldOrientation(Enum):
    ROW = auto()
    COLUMN = auto()


class Extractor:
    def __init__(self, doc_handler: BaseHandler) -> None:
        self.doc_handler = doc_handler

    def extract_closest_table(
        self,
        search_headers: List[str],
        validation_funtion: Callable[[List[str]], bool] = lambda x: True,
        minimum_proximity_ratio: float = 0,
    ) -> pd.DataFrame:
        """Extract the table in document that has the closest header to search_headers

        Args:
            search_headers (List[str]): list of itens to search in header.
            validation_funtion (Callable[[List[str]], bool], optional): function to validate if the table is valid. This function receives the table header as argument and must return True if the table is valid. Defaults to lambda x: True.
            minimum_proximity_ratio (float, optional): minimum proximity ratio to consider there is a match in header. Value must be between 0 and 100. Defaults to 0.

        Returns:
            pd.DataFrame: best match
        """
        if minimum_proximity_ratio < 0 or minimum_proximity_ratio > 100:
            raise ValueError("minimum_proximity_ratio must be between 0 and 100")

        tables = self.doc_handler.tables
        ratios = []
        for df in tables:
            if validation_funtion(df.columns.to_list()):
                ratio = self.headers_proximity_ratio(
                    document_headers=df.columns.to_list(), search_headers=search_headers
                )

                ratios.append(ratio)
            else:
                ratios.append(0)

        if len(ratios) == 0:
            return pd.DataFrame()

        best_ratio = np.max(ratios)
        if best_ratio < minimum_proximity_ratio:
            return pd.DataFrame()

        best_match = tables[np.argmax(ratios)]

        df = self.get_columns_fuzzy(best_match, search_headers)

        return df

    def extract_single_field(
        self,
        field: str,
        orientation: FieldOrientation,
        regex: List[str] = [""],
        title_regex: List[str] = [""],
        return_multiple: bool = False,
    ) -> str:
        """Extract single field of a word document based on a input string.
        The data will be extracted from tables in document

        Args:
            field (str): search field
            orientation (FieldOrientation): orientation to search the content of field
            regex (List[str], optional): list of regex to apply to content. To be a valid content there must be at least one match of regex in list. Defaults to [''].
            title_regex (List[str], optional): list of regex to apply to title. To be a valid title there must be at least one match of regex in list. Defaults to [''].
            return_multiple (bool, optional): if True, will return all matches that has the same proximity ratio. Defaults to False.

        Returns:
            str: best match
        """
        df = self.doc_handler.dictionary

        df = df[df["orientation"] == orientation.name.lower()]
        df = df[df["content"].apply(lambda x: match_regex_list(x, regex))]
        df = df[df["title"].apply(lambda x: match_regex_list(x, title_regex))]

        if df.empty:
            return ""

        df["ratio"] = df["title"].apply(lambda x: str_comparison(x, field))
        df.sort_values(by="ratio", inplace=True, ascending=False)

        try:
            if return_multiple:
                max_ratio = df["ratio"].max()
                values = df[df["ratio"] == max_ratio]["content"].to_list()
                best_match = ", ".join(values)
            else:
                best_match = df["content"].values[0]
        except IndexError:
            best_match = ""

        return best_match

    @staticmethod
    def headers_proximity_ratio(
        document_headers: List[str], search_headers: List[str]
    ) -> int:
        """Calculates a proximity ratio of two headers

        Args:
            document_headers (List[str]): headers in document
            search_headers (List[str]): search headers

        Returns:
            int: proximity ratio
        """
        matches = Extractor.headers_association(document_headers, search_headers)

        if len(matches) == 0:
            return 0

        scores = [x.score for x in matches]

        return min(scores)

    @staticmethod
    def headers_association(
        document_headers: List[str], search_headers: List[str]
    ) -> List[TableMatch]:
        # TODO I think this can be improved
        """Determine the best association of two headers

        Args:
            document_headers (List[str]): headers in document
            search_headers (List[str]): search headers

        Returns:
            List[TableMatch]: list of table headers matches
        """
        if len(search_headers) > len(document_headers):
            return []

        matches = []

        for s_header in search_headers:
            scores = [str_comparison(x, s_header) for x in document_headers]

            max_index = np.argmax(scores)
            max_score = np.max(scores)

            entry = TableMatch(
                search_term=s_header,
                original_term=document_headers[max_index],
                score=max_score,
            )

            matches.append(entry)

            document_headers.pop(max_index)

        return matches

    @staticmethod
    def get_columns_fuzzy(
        df: pd.DataFrame, columns: List[str], threshold=0
    ) -> pd.DataFrame:
        """Get columns that hat the closest match with supplied columns names
        The columns will be renamed to match the closest column name

        Args:
            df (pd.DataFrame): dataframe to search columns
            columns (List[str]): columns to search
            threshold (int, optional): minimum score to consider a match. Defaults to 0.

        Returns:
            List[str]: columns that match
        """
        association = Extractor.headers_association(df.columns.to_list(), columns)

        association = [x for x in association if x.score > threshold]

        original = [x.original_term for x in association]
        df = df[original]

        rename_dict = {x.original_term: x.search_term for x in association}
        df.rename(columns=rename_dict, inplace=True)

        return df
