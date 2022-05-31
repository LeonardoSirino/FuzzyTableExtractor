from ..handlers.docx_handler import DocxHandler
from ..extractor import Extractor, FieldOrientation

from pathlib import Path


def test_splited_tables():
    path = r"sample_docs\E004 - splited tables.docx"
    handler = DocxHandler(Path(path))
    extractor = Extractor(handler)
    df = extractor.extract_closest_table(["id", "name", "age"])
    assert len(df) == 4


def test_header_rename():
    columns = ["id", "name", "age"]

    path = r"sample_docs\E003 - typos.docx"
    handler = DocxHandler(Path(path))
    extractor = Extractor(handler)

    df = extractor.extract_closest_table(columns)
    assert df.columns.tolist() == columns


def test_extract_single_field():
    path = r"sample_docs\E005 - extract single field.docx"

    handler = DocxHandler(Path(path))
    extractor = Extractor(handler)

    name = extractor.extract_single_field("name", FieldOrientation.ROW)

    assert name == "Curitiba"


def test_no_match_table():
    path = r"sample_docs\E001 - basic content.docx"

    handler = DocxHandler(Path(path))
    extractor = Extractor(handler)

    df = extractor.extract_closest_table(["store", " game"], minimum_proximity_ratio=90)

    assert df.empty
