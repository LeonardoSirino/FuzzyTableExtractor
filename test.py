from fuzzy_table_extractor.doc_handlers import DocxHandler
from fuzzy_table_extractor.extractor import Extractor, FieldOrientation
from pathlib import Path

import logging

logging.basicConfig(level=logging.DEBUG)
logging.log(logging.INFO, "Starting test")

file_path = Path("sample_docs\sample_01.docx")
handler = DocxHandler(file_path)

extractor = Extractor(handler)
df = extractor.extract_closest_table(["ids", "name", "age"])
print(df)

johns_age = extractor.extract_single_field("John", FieldOrientation.ROW)
print(f"Johns age: {johns_age}")
