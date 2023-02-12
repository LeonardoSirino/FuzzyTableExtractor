from pathlib import Path

from .docx_handler import DocxHandler
from ..matcher import FieldOrientation

BASIC_DOC_PATH = r"src\fte\sample_docs\E001 - basic content.docx"


def test_docx_handler_tables():
    file_path = Path(BASIC_DOC_PATH)
    handler = DocxHandler(file_path)
    assert len(handler.get_tables()) == 2


def test_docx_dict():
    file_path = Path(BASIC_DOC_PATH)
    handler = DocxHandler(file_path)
    assert len(handler.get_mapping(orientation=FieldOrientation.ROW)) > 0


def test_doc_conversion():
    file_path = r"src\fte\sample_docs\E001 - basic content.doc"

    handler = DocxHandler(Path(file_path))
    assert len(handler.get_tables()) == 2
