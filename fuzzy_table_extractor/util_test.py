from .util import table_to_dataframe
from collections.abc import Sequence, MutableSequence
from dataclasses import dataclass

import pandas as pd


@dataclass
class Cell:
    text: str


@dataclass
class Row:
    cells: Sequence[Cell]


@dataclass
class Table:
    rows: Sequence[Row]


def _create_table(headers: list[str], data: list[list[str]]) -> Table:
    lines = [headers] + data
    rows: MutableSequence[Row] = []
    for line in lines:
        rows.append(Row(cells=[Cell(text=v) for v in line]))

    return Table(rows=rows)


def test_simple_table_to_dataframe():
    df = table_to_dataframe(_create_table(headers=["col1", "col2"], data=[["A", "B"]]))
    expected_df = pd.DataFrame(columns=["col1", "col2"], data=[["A", "B"]])

    diff = expected_df.compare(df)
    assert diff.empty


def test_table_no_data_to_dataframe():
    headers = ["col1", "col2"]
    table = Table(rows=[Row(cells=[Cell(text=v) for v in headers])])
    df = table_to_dataframe(table)
    expected_df = pd.DataFrame(columns=headers)

    diff = expected_df.compare(df)
    assert diff.empty


def test_table_extra_columns_to_dataframe():
    df = table_to_dataframe(
        _create_table(headers=["col1", "col2", "col3"], data=[["A", "B"]])
    )
    expected_df = pd.DataFrame(columns=["col1", "col2"], data=[["A", "B"]])

    diff = expected_df.compare(df)
    assert diff.empty


def test_table_non_rect_to_dataframe():
    df = table_to_dataframe(
        _create_table(
            headers=["col1", "col2", "col3", "col4"], data=[["A", "B"], ["C", "D", "E"]]
        )
    )
    expected_df = pd.DataFrame(
        columns=["col1", "col2", "col3"], data=[["A", "B", None], ["C", "D", "E"]]
    )

    diff = expected_df.compare(df)
    assert diff.empty


def test_table_less_columns_to_dataframe():
    df = table_to_dataframe(_create_table(headers=["col1"], data=[["A", "B"]]))
    expected_df = pd.DataFrame(columns=["col1", "unnamed col 0"], data=[["A", "B"]])

    diff = expected_df.compare(df)
    assert diff.empty
