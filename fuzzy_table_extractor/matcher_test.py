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
