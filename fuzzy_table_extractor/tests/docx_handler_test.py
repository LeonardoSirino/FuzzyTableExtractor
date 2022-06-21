from pathlib import Path

from ..handlers.docx_handler import DocxHandler, TreeDocxHandler

BASIC_DOC_PATH = r"sample_docs\E001 - basic content.docx"


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
    file_path = r"sample_docs\E001 - basic content.doc"

    handler = DocxHandler(Path(file_path))
    assert len(handler.words) > 0


def test_tree_doc_handler():
    file_path = r"sample_docs\E006 - doc with sections.docx"

    handler = TreeDocxHandler(Path(file_path))
    root = handler.root

    assert root is not None
    assert len(root.get_paragraphs(recursive=True)) > 0

    first_section = root.nodes[0]
    assert len(first_section.get_tables(recursive=True)) == 1
