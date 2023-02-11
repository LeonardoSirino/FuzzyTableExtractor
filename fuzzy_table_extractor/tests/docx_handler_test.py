from pathlib import Path

from ..handlers.docx_handler import DocxHandler

BASIC_DOC_PATH = r"src\fte\sample_docs\E001 - basic content.docx"


def test_docx_handler_words():
    file_path = Path(BASIC_DOC_PATH)
    handler = DocxHandler(file_path)
    assert len(handler.words) > 0


def test_docx_handler_tables():
    file_path = Path(BASIC_DOC_PATH)
    handler = DocxHandler(file_path)
    assert len(handler.tables) == 2


def test_docx_dict():
    file_path = Path(BASIC_DOC_PATH)
    handler = DocxHandler(file_path)
    assert len(handler.dictionary) > 0


def test_doc_conversion():
    file_path = r"src\fte\sample_docs\E001 - basic content.doc"

    handler = DocxHandler(Path(file_path))
    assert len(handler.words) > 0
