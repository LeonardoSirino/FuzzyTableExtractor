# Fuzzy table extractor
## Introduction
This project aims to help data extraction from unstructured sources, like Word and pdf files, web documents, and so on.

The library has 2 main components: the file handler, which is responsible for identifying tables in the document and returning these in a more controlled way; the extractor, which searches in document's tables and returns the one with the highest proximity, using for this a fuzzy string comparison algorithm.

Currently, there is only a handler for Docx files, but in the future, this will be expanded to other sources.

## Using the library
The usage of the library is very simple: first, a handler for the file must be created, then this object is used to create an instance of Extractor, which will contain methods for data extraction.

Here is an example of table extraction for a very simple document:

```python
from pathlib import Path

from fuzzy_table_extractor.doc_handlers import DocxHandler
from fuzzy_table_extractor.extractor import Extractor

file_path = Path(r"sample_docs\E001 - basic content.docx")

handler = DocxHandler(file_path)
extractor = Extractor(handler)

df = extractor.extract_closest_table(["id", "name", "age"])
print(df)
```
For a document that looks like this:

![some image](assets\basic_document.png)

The ouput is:
```
  id  name age
0  0  Paul  25
1  1  John  32
```