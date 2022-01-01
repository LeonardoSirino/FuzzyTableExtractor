from fuzzy_table_extractor.doc_handlers import DocxHandler
from fuzzy_table_extractor.extractor import Extractor, FieldOrientation
from pathlib import Path

import logging

logging.basicConfig(level=logging.DEBUG)
logging.log(logging.INFO, "Starting test")

path = r"sample_docs\E004 - splited tables.docx"

file_path = Path(path)
handler = DocxHandler(file_path)

extractor = Extractor(handler)
df = extractor.extract_closest_table(["id", "name", "age"])
print(df)
