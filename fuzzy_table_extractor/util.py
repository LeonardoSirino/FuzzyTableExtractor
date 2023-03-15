import functools
from collections.abc import Sequence
from typing import Protocol

import pandas as pd
from fuzzywuzzy import fuzz
from unidecode import unidecode


class _Cell(Protocol):
    text: str


class _Row(Protocol):
    cells: Sequence[_Cell]


class _Table(Protocol):
    rows: Sequence[_Row]


def table_to_dataframe(table: _Table) -> pd.DataFrame:
    """Converts a docs Table to a pandas Dataframe.

    If headers is greater than the data, the extra column names will be removed. If there
    are less column names, `unnamed col` names will be created to fill as needed.

    Args:
        table (_Table): Table object found in the .docx document.

    Returns:
        pd.DataFrame: Dataframe holding the data in the table.
    """
    headers = [cell.text for cell in table.rows[0].cells]
    if len(table.rows) == 1:
        return pd.DataFrame(columns=headers)

    data = [[cell.text for cell in row.cells] for row in table.rows[1:]]
    max_size = max(len(row) for row in data)

    if len(headers) > max_size:
        headers = headers[:max_size]
    elif max_size > len(headers):
        size_diff = max_size - len(headers)
        headers += [f"unnamed col {i}" for i in range(size_diff)]

    df = pd.DataFrame(columns=headers, data=data)

    return df


@functools.cache
def str_comparison(text_a: str, text_b: str) -> int:
    """Get the proximity ratio of 2 strings

    Args:
        text_a (str): input text A
        text_b (str): input text B

    Returns:
        int: ration of proximity for 2 inputs [0 - 100]
    """

    if not isinstance(text_a, str) or not isinstance(text_b, str):
        return 0

    a = text_a.replace(" ", "").replace("\n", "")
    b = text_b.replace(" ", "").replace("\n", "")

    a = unidecode(a.lower())
    b = unidecode(b.lower())

    return fuzz.partial_ratio(a, b)
