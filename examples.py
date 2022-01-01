from pathlib import Path

from fuzzy_table_extractor.doc_handlers import DocxHandler
from fuzzy_table_extractor.extractor import Extractor, FieldOrientation


def get_basic_table():
    path = r"sample_docs\E001 - basic content.docx"

    file_path = Path(path)
    handler = DocxHandler(file_path)

    extractor = Extractor(handler)
    df = extractor.extract_closest_table(["id", "name", "age"])
    print(df)


if __name__ == "__main__":
    get_basic_table()
