Basic usage
===========

Getting a table from a Microsoft Word documentation
----------------------------------------------------

.. code-block:: python

   from pathlib import Path

   from fuzzy_table_extractor.doc_handlers import DocxHandler
   from fuzzy_table_extractor.extractor import Extractor, FieldOrientation

   path = r"path_to_document.docx"

   file_path = Path(path)
   handler = DocxHandler(file_path)

   extractor = Extractor(handler)
   df = extractor.extract_closest_table(["id", "name", "age"])
   print("This is the result extraction of a very simple document:")
   print(df)