from pathlib import Path
from typing import Protocol, Sequence

import pandas as pd
import pytest

from ..matcher import FieldOrientation
from .docx_handler import DocxHandler
from .docx_xml_handler import DocxXMLHandler

BASIC_DOC_PATH = r"src\fte\sample_docs\E001 - basic content.docx"


class Handler(Protocol):
    def get_tables(self) -> Sequence[pd.DataFrame]:
        ...

    def get_mapping(self, orientation: FieldOrientation) -> dict[str, Sequence[str]]:
        ...


_HANDLERS_TO_TEST = [DocxHandler(Path(BASIC_DOC_PATH)), DocxXMLHandler(BASIC_DOC_PATH)]


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


def test_doc_conversion():
    file_path = r"src\fte\sample_docs\E001 - basic content.doc"

    handler = DocxHandler(Path(file_path))
    assert len(handler.get_tables()) == 2
