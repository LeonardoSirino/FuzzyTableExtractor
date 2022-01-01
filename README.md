# Fuzzy table extractor
## Introduction
This project aims to help data extraction from unstructured sources, like Word and pdf files, web documents, and so on.

The library has 2 main components: the file handler, which is responsible for identifying tables in the document and returning these in a more standarlized way; the extractor, which searches in document's tables and returns the one with the highest proximity, using for this a fuzzy string comparison algorithm.

Currently, there is only a handler for Docx files, but in the future, this will be expanded to other sources.

## Installation
The library is available on PyPI:
```
pip install fuzzy-table-extractor
```

## Using the library
### Extracting tables
The usage of the library is very simple: first, a handler for the file must be created, then this object is used to create an instance of Extractor, which will contain methods for data extraction.

Here is an example of table extraction for a very simple document:

```python
from pathlib import Path

from fuzzy_table_extractor.doc_handlers import DocxHandler
from fuzzy_table_extractor.extractor import Extractor

file_path = Path(r"path_to_document.docx")

handler = DocxHandler(file_path)
extractor = Extractor(handler)

df = extractor.extract_closest_table(search_headers=["id", "name", "age"])
print(df)
```
For a document that looks like this:

![Basic document](https://github.com/LeonardoSirino/FuzzyTableExtractor/blob/main/assets/basic_document.png?raw=true)

The output is:
```
  id  name age
0  0  Paul  25
1  1  John  32
```

Due to the fuzzy match used to select the closest table, this library is resilient to typos. As an example, using the same code above, but now for a document like this:

![Typos in document](assets\typos_in_document.png)
The output is:
```
  id  name age
0  0  Paul  25
1  1  John  32
2  2   Bob  56
```
### Extracting single field
There is also the possibility to extract only a single field (cell) from a document. Here is an example of how to do this with the library:

```python
from pathlib import Path

from fuzzy_table_extractor.doc_handlers import DocxHandler
from fuzzy_table_extractor.extractor import Extractor, FieldOrientation

file_path = Path(r"path_to_document.docx")

handler = DocxHandler(file_path)
extractor = Extractor(handler)

area = extractor.extract_single_field(field="area", 
                                      orientation=FieldOrientation.ROW)
print(area)
```

For a document like this:
![Extracting single field](assets\extract_single_field.png)

The output is:
```
430.9 km2
```

The file [examples.py](https://github.com/LeonardoSirino/FuzzyTableExtractor/blob/main/examples.py) contains other examples of how to use the library


## TODO
- [ ] Add to README a guide on how to contribute to project
- [ ] Expand test coverage
- [ ] Create a handler for pdf files