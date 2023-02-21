import functools
import xml.etree.ElementTree
import zipfile
from collections import defaultdict, deque
from typing import MutableSequence, Sequence
from xml.etree.ElementTree import Element

import numpy as np
import pandas as pd

from ..matcher import FieldOrientation

_WORD_NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_TEXT = _WORD_NAMESPACE + "t"
_TABLE = _WORD_NAMESPACE + "tbl"
_ROW = _WORD_NAMESPACE + "tr"
_CELL = _WORD_NAMESPACE + "tc"


class DocxXMLHandler:
    def __init__(self, file_path: str) -> None:
        self._file_path = file_path

    @functools.cache
    def get_tables(self) -> Sequence[pd.DataFrame]:
        with zipfile.ZipFile(self._file_path) as docx:
            tree = xml.etree.ElementTree.XML(docx.read("word/document.xml"))

        tables: MutableSequence[pd.DataFrame] = []
        for table in tree.iter(_TABLE):
            data: MutableSequence[MutableSequence[str]] = []
            for row in table.findall(f"./{_ROW}"):
                data.append(
                    [_get_combined_text(cell) for cell in row.findall(f"./{_CELL}")]
                )

            df = pd.DataFrame(data)
            df.columns = df.iloc[0]
            df = df[1:]

            tables.append(df)

        return tables

    @functools.cached_property
    def _mapping(self) -> dict[FieldOrientation, dict[str, Sequence[str]]]:
        mapping: dict[FieldOrientation, dict[str, MutableSequence[str]]] = {
            FieldOrientation.ROW: defaultdict(list),
            FieldOrientation.COLUMN: defaultdict(list),
        }
        for df in self.get_tables():
            records = _table_to_records(df)
            for row in records:
                for i, key in enumerate(row[:-1]):
                    mapping[FieldOrientation.ROW][key].append(row[i + 1])

            for col in np.array(records).T.tolist():
                for i, key in enumerate(col[:-1]):
                    mapping[FieldOrientation.COLUMN][key].append(col[i + 1])

        return mapping

    def get_mapping(self, orientation: FieldOrientation) -> dict[str, Sequence[str]]:
        """Retrieves the mapping of values in the document.

        Args:
            orientation (FieldOrientation): Which direction to get the mapping.

        Returns:
            dict[str, Sequence[str]]: A mapping of key names to all contents in the
                document.
        """
        return self._mapping[orientation]


def _table_to_records(df: pd.DataFrame) -> Sequence[Sequence[str]]:
    return [df.columns.to_numpy(dtype=str).tolist()] + df.to_numpy(dtype=str).tolist()


def _get_combined_text(cell: Element) -> str:
    fragments: MutableSequence[str] = []
    q = deque(cell.findall("./*"))
    while q:
        cell = q.popleft()
        if cell.tag == _TABLE:
            continue

        if cell.tag == _TEXT:
            fragments.append(str(cell.text))

        q.extendleft(cell.findall("./*")[::-1])

    return "\n".join(fragments)
