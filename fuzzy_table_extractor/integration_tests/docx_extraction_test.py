from pathlib import Path

from ..handlers.docx_handler import DocxHandler
from ..matcher import FieldOrientation, Matcher


def test_splited_tables():
    path = Path(r"src\fte\sample_docs\E001 - basic content.doc").resolve()
    handler = DocxHandler(path)
    matcher = Matcher(handler)
    df, _ = matcher.match_table(["id", "name", "age"])
    assert len(df) == 2


def test_header_rename():
    columns = ["id", "name", "age"]

    path = Path(r"src\fte\sample_docs\E003 - typos.docx").resolve()
    handler = DocxHandler(path)
    matcher = Matcher(handler)

    df, _ = matcher.match_table(columns)
    assert df.columns.tolist() == columns


def test_extract_single_field():
    path = Path(r"src\fte\sample_docs\E005 - extract single field.docx").resolve()

    handler = DocxHandler(path)
    matcher = Matcher(handler)

    name, _ = matcher.match_field("name", FieldOrientation.ROW)

    assert name == "Curitiba"


def test_low_score_match():
    path = Path(r"src\fte\sample_docs\E001 - basic content.docx").resolve()

    handler = DocxHandler(path)
    matcher = Matcher(handler)

    _, score = matcher.match_table(["store", " game"])

    assert score < 90
