from collections.abc import Sequence

import pandas as pd

from .matcher import FieldOrientation, Matcher


class FakeHandler:
    def __init__(
        self, mapping: dict[str, Sequence[str]], tables: Sequence[pd.DataFrame]
    ) -> None:
        self._mapping = mapping
        self._tables = tables

    def get_tables(self) -> Sequence[pd.DataFrame]:
        return self._tables

    def get_mapping(self, orientation: FieldOrientation) -> dict[str, Sequence[str]]:
        return self._mapping


def test_match_field():
    matcher = Matcher(
        FakeHandler(mapping={"title": ["content", "other_content"]}, tables=[])
    )

    field, _ = matcher.match_field(
        "title", orientation=FieldOrientation.ROW, return_multiple=False
    )
    assert field == "content"


def test_match_field_return_multiple():
    matcher = Matcher(
        FakeHandler(mapping={"title": ["content", "other_content"]}, tables=[])
    )

    field, _ = matcher.match_field(
        "title", orientation=FieldOrientation.ROW, return_multiple=True
    )
    assert field == "content, other_content"


def test_match_field_title_regex():
    matcher = Matcher(
        FakeHandler(
            mapping={"title": ["content", "other_content"], "other": ["other"]}, tables=[]
        )
    )

    field, _ = matcher.match_field(
        "title",
        orientation=FieldOrientation.ROW,
        return_multiple=False,
        title_regex=["other"],
    )
    assert field == "other"


def test_match_field_content_regex():
    matcher = Matcher(
        FakeHandler(
            mapping={"title": ["content", "other_content"], "other": ["other"]}, tables=[]
        )
    )

    field, _ = matcher.match_field(
        "title",
        orientation=FieldOrientation.ROW,
        return_multiple=False,
        regex=["other"],
    )
    assert field == "other_content"


def test_match_table():
    matcher = Matcher(
        FakeHandler(
            mapping={},
            tables=_create_fake_dataframes([["nothing", "other"], ["title1", "name1"]]),
        )
    )

    table, _ = matcher.match_table(search_headers=["title", "name"], rename_columns=False)
    assert table.columns.to_list() == ["title1", "name1"]


def test_match_table_renaming():
    matcher = Matcher(
        FakeHandler(
            mapping={},
            tables=_create_fake_dataframes([["nothing", "other"], ["title1", "name1"]]),
        )
    )

    table, _ = matcher.match_table(search_headers=["title", "name"], rename_columns=True)
    assert table.columns.to_list() == ["title", "name"]


def test_match_table_out_of_order():
    matcher = Matcher(
        FakeHandler(
            mapping={},
            tables=_create_fake_dataframes([["nothing", "other"], ["name", "title"]]),
        )
    )

    table, _ = matcher.match_table(search_headers=["title", "name"])
    assert table.columns.to_list() == ["title", "name"]


def test_match_table_more_headers():
    matcher = Matcher(
        FakeHandler(
            mapping={},
            tables=_create_fake_dataframes(
                [["nothing", "other", "title"], ["name", "title", "other", "non related"]]
            ),
        )
    )

    table, _ = matcher.match_table(search_headers=["title", "name"])
    assert table.columns.to_list() == ["title", "name"]


def test_match_table_duplicated_columns():
    matcher = Matcher(
        FakeHandler(
            mapping={},
            tables=_create_fake_dataframes(
                [
                    ["nothing", "other", "title"],
                    ["name", "title", "title", "non related"],
                ]
            ),
        )
    )

    table, _ = matcher.match_table(search_headers=["title", "name"])
    assert table.columns.to_list() == ["title", "name"]


def _create_fake_dataframes(headers: Sequence[Sequence[str]]) -> Sequence[pd.DataFrame]:
    return [pd.DataFrame(columns=h) for h in headers]
