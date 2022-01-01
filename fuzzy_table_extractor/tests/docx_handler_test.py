from ..doc_handlers import DocxHandler
from pathlib import Path


def test_docx_handler_words():
    file_path = Path(r"sample_docs\sample_01.docx")
    handler = DocxHandler(file_path)
    assert len(handler.words) > 0


def test_docx_handler_tables():
    file_path = Path(r"sample_docs\sample_01.docx")
    handler = DocxHandler(file_path)
    assert len(handler.tables) > 0


def test_docx_dict():
    file_path = Path(r"sample_docs\sample_01.docx")
    handler = DocxHandler(file_path)
    assert len(handler.dictionary) > 0
