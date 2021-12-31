from fuzzy_table_extractor.doc_handlers import BaseHandler, DocxHandler
from pathlib import Path

import logging

logging.basicConfig(level=logging.DEBUG)
logging.log(logging.INFO, "Starting test")

file_path = Path("sample_docs\sample_01.docx")
handler = DocxHandler(file_path)

tables = handler.tables[0]
print(tables)
