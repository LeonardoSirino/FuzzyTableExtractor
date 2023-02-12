from ..handlers.docx_handler import DocxHandler
from ..matcher import Extractor, FieldOrientation

from pathlib import Path


def test_splited_tables():
    path = Path(r"src\fte\sample_docs\E001 - basic content.doc").resolve()
    handler = DocxHandler(path)
    extractor = Extractor(handler)
    df = extractor.extract_closest_table(["id", "name", "age"])
    assert len(df) == 2


def test_header_rename():
    columns = ["id", "name", "age"]

    path = Path(r"src\fte\sample_docs\E003 - typos.docx").resolve()
    handler = DocxHandler(path)
    extractor = Extractor(handler)

    df = extractor.extract_closest_table(columns)
    assert df.columns.tolist() == columns


def test_extract_single_field():
    path = Path(r"src\fte\sample_docs\E005 - extract single field.docx").resolve()

    handler = DocxHandler(path)
    extractor = Extractor(handler)

    name = extractor.extract_single_field("name", FieldOrientation.ROW)

    assert name == "Curitiba"


def test_no_match_table():
    path = Path(r"src\fte\sample_docs\E001 - basic content.docx").resolve()

    handler = DocxHandler(path)
    extractor = Extractor(handler)

    df = extractor.extract_closest_table(["store", " game"], minimum_proximity_ratio=90)

    assert df.empty
