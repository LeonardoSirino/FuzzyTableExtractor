from ..doc_handlers import DocxHandler
from ..extractor import Extractor

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
