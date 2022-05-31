from pathlib import Path

from fuzzy_table_extractor.handlers.docx_handler import DocxHandler
from fuzzy_table_extractor.extractor import Extractor, FieldOrientation


def get_basic_table():
    path = r"sample_docs\E001 - basic content.docx"

    file_path = Path(path)
    handler = DocxHandler(file_path)

    extractor = Extractor(handler)
    df = extractor.extract_closest_table(["id", "name", "age"])
    print("This is the result extraction of a very simple document:")
    print(df)
    print("\n")


def getting_in_inner_tables():
    path = r"sample_docs\E002 - inner tables.docx"

    file_path = Path(path)
    handler = DocxHandler(file_path)

    extractor = Extractor(handler)
    df = extractor.extract_closest_table(["id", "city", "population"])
    print(
        "This is the result extraction of a document that has the target table inside other table:"
    )
    print(df)
    print("\n")


def document_with_typos():
    path = r"sample_docs\E003 - typos.docx"

    file_path = Path(path)
    handler = DocxHandler(file_path)

    extractor = Extractor(handler)
    df = extractor.extract_closest_table(["id", "name", "age"])
    print("This is the result extraction of a document that has typos in the table:")
    print(df)
    print("\n")


def splited_tables():
    path = r"sample_docs\E004 - splited tables.docx"

    file_path = Path(path)
    handler = DocxHandler(file_path)

    extractor = Extractor(handler)
    df = extractor.extract_closest_table(["id", "name", "age"])
    print(
        "This is the result extraction of a document that has 2 tables with the same header:"
    )
    print(df)
    print("\n")


def getting_a_field():
    path = r"sample_docs\E005 - extract single field.docx"

    file_path = Path(path)
    handler = DocxHandler(file_path)

    extractor = Extractor(handler)
    df = extractor.extract_single_field(field="area", orientation=FieldOrientation.ROW)
    print("This is the result extraction of a single field oriented in row:")
    print(df)
    print("\n")


if __name__ == "__main__":
    get_basic_table()
    getting_in_inner_tables()
    document_with_typos()
    splited_tables()
    getting_a_field()
