from pathlib import Path
from typing import Protocol, Sequence

import pandas as pd
import pytest

from ..matcher import FieldOrientation
from .docx_handler import DocxHandler, DocxXMLHandler

_BASIC_DOC_PATH = r"src\fte\sample_docs\E001 - basic content.docx"


class Handler(Protocol):
    def get_tables(self) -> Sequence[pd.DataFrame]:
        ...

    def get_mapping(self, orientation: FieldOrientation) -> dict[str, Sequence[str]]:
        ...


_HANDLERS_TO_TEST = [DocxHandler(Path(_BASIC_DOC_PATH)), DocxXMLHandler(_BASIC_DOC_PATH)]


@pytest.mark.parametrize("handler", _HANDLERS_TO_TEST)
def test_docx_handler_identify_tables(handler: Handler):
    assert len(handler.get_tables()) == 2


@pytest.mark.parametrize("handler", _HANDLERS_TO_TEST)
def test_docx_handler_assert_headers(handler: Handler):
    assert handler.get_tables()[0].columns.to_list() == ["ID", "City", "State"]
    assert handler.get_tables()[1].columns.to_list() == [
        "ID",
        "Initials",
        "Name",
        "Country",
        "Age",
    ]


@pytest.mark.parametrize("handler", _HANDLERS_TO_TEST)
def test_docx_handler_identify_mappings(handler: Handler):
    assert len(handler.get_mapping(orientation=FieldOrientation.ROW)) > 0


@pytest.mark.parametrize("handler", _HANDLERS_TO_TEST)
def test_docx_handler_assert_mapping_values(handler: Handler):
    assert handler.get_mapping(orientation=FieldOrientation.COLUMN)["City"] == [
        "Curitiba"
    ]
    assert handler.get_mapping(orientation=FieldOrientation.ROW)["Blumenau"] == [
        "Santa Catarina"
    ]


_DOC_FILE_PATH = r"src\fte\sample_docs\E001 - basic content.doc"


@pytest.mark.parametrize(
    "handler",
    [
        DocxHandler(Path(_DOC_FILE_PATH)),
        DocxXMLHandler(str(Path(_DOC_FILE_PATH).resolve())),
    ],
)
def test_doc_conversion(handler: Handler):
    assert len(handler.get_tables()) == 2
