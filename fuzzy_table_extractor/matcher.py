import re
from collections.abc import MutableSequence, Sequence
from dataclasses import dataclass
from enum import Enum, auto
from typing import Any, Callable, Optional, Protocol

import numpy as np
import pandas as pd
from scipy import optimize

from .util import str_comparison


class NoValidMatchError(Exception):
    def __init__(self) -> None:
        super().__init__("No valid matches found.")


@dataclass
class _TableMatch:
    search_term: str
    original_term: str
    score: int


class FieldOrientation(Enum):
    ROW = auto()
    COLUMN = auto()


class Handler(Protocol):
    def get_tables(self) -> Sequence[pd.DataFrame]:
        ...

    def get_mapping(self, orientation: FieldOrientation) -> dict[str, Sequence[str]]:
        ...


class Matcher:
    def __init__(self, handler: Handler) -> None:
        self._handler = handler

    def match_table(
        self,
        search_headers: Sequence[str],
        validation_funtion: Optional[Callable[[Sequence[Any]], bool]] = None,
        rename_columns: bool = True,
    ) -> tuple[pd.DataFrame, int]:
        """Matches a table supplied by the handler.

        Args:
            search_headers (Sequence[str]): Sequence of column names to search for tables
            supplied by the handler.
            validation_funtion (Optional[Callable[[Sequence[Any]], bool]], optional):
                Function to validate headers supplied by the handler. This function
                receives the headers from a table and should return whether the table is
                valid for matching. Defaults to None, which means that no headers
                validation will be applied.
            rename_columns (bool, optional): Whether to rename columns to match search
                headers. Defaults to True.

        Returns:
            tuple[pd.DataFrame, int]: Tuple with:
                - Table with best match for search_headers;
                - Score of the match.
        """
        tables = self._handler.get_tables()
        if validation_funtion is not None:
            tables = [t for t in tables if validation_funtion(t.columns.to_list())]

        if not tables:
            raise NoValidMatchError

        ratios: MutableSequence[int] = []
        for t in tables:
            headers = [str(h) for h in t.columns.to_list()]
            ratios.append(
                sequence_proximity_ratio(ref_seq=headers, test_seq=search_headers)
            )

        df = tables[np.argmax(ratios)]
        if rename_columns:
            df = get_columns_fuzzy(df, search_headers)

        return df, max(ratios)

    def match_field(
        self,
        field: str,
        orientation: FieldOrientation,
        regex: Sequence[str] = [""],
        title_regex: Sequence[str] = [""],
        return_multiple: bool = False,
    ) -> tuple[str, int]:
        """Match a field in the handler mapping.

        Args:
            field (str): Field to search in the handler mapping. The field is key name.
            orientation (FieldOrientation): Which direction to retrieve the content.
            regex (Sequence[str], optional): List of regex patterns that the content must
                match to be valid. Defaults to [""].
            title_regex (Sequence[str], optional): List of regex patterns that the key
                name must match to be valid. Defaults to [""].
            return_multiple (bool, optional): Whether to return multiple matches if they
                have the same proximity ratio. If set to True, the contents that have the
                highest ratio will be combined into a single string, delimited by commas,
                and returned as the match. Defaults to False.

        Returns:
            tuple[str, int]: The content with highest proximity ratio and the value of this ratio.
        """

        map_ = self._handler.get_mapping(orientation=orientation)

        if title_regex:
            title_pattern = re.compile("|".join(title_regex))
            map_ = {k: v for k, v in map_.items() if title_pattern.search(k)}

        if regex:
            c_pattern = re.compile("|".join(regex))
            map_ = {k: [i for i in v if c_pattern.search(i)] for k, v in map_.items()}
            map_ = {k: v for k, v in map_.items() if v}

        if not map_:
            return "", 0

        df = (
            pd.DataFrame(data=map_.items(), columns=["key", "value"])
            .explode(["value"])
            .reset_index(drop=True)
        )

        df["ratio"] = df["key"].apply(lambda x: str_comparison(x, field))
        df.sort_values(by="ratio", inplace=True, ascending=False)

        max_ratio = df["ratio"].max()

        try:
            if return_multiple:
                values = df[df["ratio"] == max_ratio]["value"].to_list()
                best_match = ", ".join(values)
            else:
                best_match = df["value"].values[0]
        except IndexError:
            return "", 0

        return best_match, max_ratio


def sequence_proximity_ratio(
    ref_seq: Sequence[str], test_seq: Sequence[str], optimal_match: bool = True
) -> int:
    """Calculates the proximity ratio for 2 sequences of strings.

    To calculate the proximity ratio, first the best association between terms in the test
    sequence and terms in the reference sequence is found. Each pair in this association
    has a proximity ratio between the terms, the proximity ratio of the sequence is the
    smallest proximity ratio between terms.

    Args:
        ref_seq (Sequence[str]): Reference sequence of strings.
        test_seq (Sequence[str]): Test sequence of strings.
        optimal_match (bool): Whether to find the optimal association between terms in the
            reference and test sequences. The optimal algorithm solves the linear sum
            assignment problem, which has O(n³) complexity, while the non optimal approach
            uses a naive algorithm with O(n²) complexity. Defaults to True.

    Returns:
        int: Proximity ratio of the sequences.
    """

    # TODO benchmark 2 approaches
    if optimal_match:
        matches = _optimal_sequence_matching(ref_seq, test_seq)
    else:
        matches = _naive_sequence_matching(ref_seq, test_seq)

    if len(matches) == 0:
        return 0

    scores = [x.score for x in matches]
    return min(scores)


def _naive_sequence_matching(
    ref_seq: Sequence[str], test_seq: Sequence[str]
) -> Sequence[_TableMatch]:
    """Finds the association between terms in 2 sequences in a fast, but naive, way."""

    if len(test_seq) > len(ref_seq):
        return []

    # creating a copy of the list to prevent modifying it inplace
    ref_seq_ = list(ref_seq[:])

    matches: Sequence[_TableMatch] = []
    for term in test_seq:
        scores = [str_comparison(x, term) for x in ref_seq_]

        max_index = np.argmax(scores)
        max_score = np.max(scores)

        entry = _TableMatch(
            search_term=term,
            original_term=ref_seq_[max_index],
            score=max_score,
        )

        matches.append(entry)

        ref_seq_.pop(max_index)

    return matches


def _optimal_sequence_matching(
    ref_seq: Sequence[str], test_seq: Sequence[str]
) -> Sequence[_TableMatch]:
    """Finds the optimal association between terms in 2 sequences."""
    cost_matrix = [[str_comparison(r, t) for t in test_seq] for r in ref_seq]

    row_ind: Sequence[int]
    col_ind: Sequence[int]
    row_ind, col_ind = optimize.linear_sum_assignment(cost_matrix, maximize=True)

    matches: Sequence[_TableMatch] = []
    for r, c in zip(row_ind, col_ind):
        matches.append(
            _TableMatch(
                search_term=test_seq[c], original_term=ref_seq[r], score=cost_matrix[r][c]
            )
        )

    return matches


def get_columns_fuzzy(
    df: pd.DataFrame, columns: Sequence[str], threshold: int = 0
) -> pd.DataFrame:
    """Gets columns in a dataframe using fuzzy matching.

    Args:
        df (pd.DataFrame): Dataframe to get columns from.
        columns (Sequence[str]): Sequence of columns to retrieve from the Dataframe.
        threshold (int, optional): Minimum proximity ratio between search and dataframe
            columns to consider a valid match. Defaults to 0, which means that any match
            is valid.

    Raises:
        NoValidMatch: If the proximity ratio between the search columns and the columns in
            the dataframe is smaller than the threshold.

    Returns:
        pd.DataFrame: Dataframe with selected columns. Note that the columns in the
            dataframe will be renamed to match values inputed in the function.
    """
    df_ = df.copy()

    association = _optimal_sequence_matching(df_.columns.to_list(), columns)
    min_ratio = min([m.score for m in association])
    if min_ratio < threshold:
        raise NoValidMatchError

    original = [x.original_term for x in association]
    df_ = df_[original]

    rename_dict = {x.original_term: x.search_term for x in association}
    df_.rename(columns=rename_dict, inplace=True)

    if df_.columns.duplicated().any():
        df_ = df_.loc[:, ~df_.columns.duplicated()]

    # ensure that columns will return in the supplied columns order.
    return df_[columns]
