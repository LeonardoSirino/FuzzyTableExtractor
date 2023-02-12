import functools
import os
from collections import defaultdict
from collections.abc import MutableSequence, Sequence
from pathlib import Path

import pandas as pd
import win32com.client as win32
from docx.api import Document
from docx.table import Table

from ..matcher import FieldOrientation
from ..util import table_to_dataframe

_MAX_NESTING_LEVEL = 10


class DocxHandler:
    """Handler for Microsoft Word Documents.

    It is suposed to use the newer .docx format for Word Documents, but if a .doc file is
    provided, it will be converted to .docx format.
    """

    def __init__(self, file_path: Path, temp_folder: Path = Path("temp")) -> None:
        """Creates a new DocxHandler instance.

        Args:
            file_path (Path): Path to file to be handled
            temp_folder (Path, optional): Path to temporary folder where docx files will
                be created when the supplied file has .doc extension. Defaults to 'temp'.
        """
        self._file_path = file_path
        self._temp_folder = temp_folder

    def get_mapping(self, orientation: FieldOrientation) -> dict[str, Sequence[str]]:
        """Retrieves the mapping of values in the document.

        Args:
            orientation (FieldOrientation): Which direction to get the mapping.

        Returns:
            dict[str, Sequence[str]]: A mapping of key names to all contents in the
                document.
        """
        return self._mapping[orientation]

    @functools.cached_property
    def _mapping(self) -> dict[FieldOrientation, dict[str, Sequence[str]]]:
        row_mapping: dict[str, MutableSequence[str]] = defaultdict(list)
        for table in self._docx_tables:
            n_cols = len(table.columns)

            for row in table.rows:
                cells = row.cells
                values = [cell.text for cell in cells]
                for k in range(n_cols - 1):
                    if (k + 1) >= len(values):
                        continue

                    title = values[k]
                    content = values[k + 1]
                    if title == content or title == "" or content == "":
                        continue

                    row_mapping[title].append(content)

        col_mapping: dict[str, MutableSequence[str]] = defaultdict(list)
        for table in self._docx_tables:
            n_rows = len(table.rows)

            for col in table.columns:
                cells = col.cells
                values = [cell.text for cell in cells]
                for k in range(n_rows - 1):
                    if (k + 1) >= len(values):
                        continue

                    title = values[k]
                    content = values[k + 1]
                    if title == content or title == "" or content == "":
                        continue

                    col_mapping[title].append(content)

        return {
            FieldOrientation.ROW: row_mapping,
            FieldOrientation.COLUMN: col_mapping,
        }

    @functools.cache
    def get_tables(self) -> Sequence[pd.DataFrame]:
        """Gets all tables in the document.

        Returns:
            Sequence[pd.DataFrame]: Sequence of tables in the document.
        """
        tables = self._docx_tables

        # Getting subset of tables that has a merged header
        splited_tables = []
        aux_table = []
        for k, table in enumerate(tables):
            for row in table.rows:
                values = [x.text for x in row.cells]
                if len(set(values)) == 1:
                    if len(values[0]) < 200:
                        # this is a merged table header with limited text lenght
                        splited_tables.append(aux_table)
                        aux_table = []
                else:
                    aux_table.append(values)

        dfs = []

        # Gettind data from conventional tables
        for table in tables:
            df = table_to_dataframe(table)

            if len(df) > 0:
                dfs.append(df)

        # Getting data from splited tables
        for table in splited_tables:
            if len(table) == 0:
                continue

            header = table[0]
            data = table[1:]
            header_size = len(header)

            aux = []

            # Removing data with different length from header
            for line in data:
                if len(line) == header_size:
                    aux.append(line)
                else:
                    break

            df = pd.DataFrame(columns=header, data=aux)
            dfs.append(df)

        dfs = self._merge_dfs(dfs)

        return dfs

    @functools.cached_property
    def _docx_tables(self) -> Sequence[Table]:
        """List of tables in docx document"""
        tables: MutableSequence[Table] = self._document.tables

        new_tables = tables[:]
        count = 0

        # getting all inner tables
        while new_tables:
            aux = []
            for table in new_tables:
                for row in table.rows:
                    for cell in row.cells:
                        aux.extend(cell.tables)

            tables.extend(aux)
            new_tables = aux

            count += 1
            if count > _MAX_NESTING_LEVEL:
                break

        return tables

    @functools.cached_property
    def _document(self) -> Document:
        """Open document and creates a docx file if necessary.

        Returns:
            Document: Word document object.
        """
        file_path = self._file_path
        if file_path.suffix[1:] == "doc":
            file_path = self._create_docx_file()

        return Document(str(file_path.resolve()))

    @property
    def _docx_file_path(self) -> Path:
        if self._file_path.suffix == ".docx":
            return self._file_path

        return self._temp_folder / f"_aux_{self._file_path.stem}_file.docx"

    def _create_docx_file(self) -> Path:
        docx_file_path = self._docx_file_path

        if os.path.exists(docx_file_path):
            return docx_file_path

        self._temp_folder.mkdir(exist_ok=True, parents=True)

        # NOTE
        # If errors are found, do this
        # clear contents of C:\Users\<username>\AppData\Local\Temp\gen_py

        word = win32.gencache.EnsureDispatch("Word.Application")

        doc = word.Documents.Open(str(self._file_path.resolve()))
        doc.Activate()

        word.ActiveDocument.SaveAs(
            str(docx_file_path.resolve()),
            FileFormat=win32.constants.wdFormatXMLDocument,
        )
        doc.Close(False)

        return docx_file_path

    def _merge_dfs(self, dfs: Sequence[pd.DataFrame]) -> Sequence[pd.DataFrame]:
        """Merge dataframes that has the same header and drop duplicated lines.

        Args:
            dfs (Sequence[pd.DataFrame]): Sequence of dataframes from doc extraction.

        Returns:
            Sequence[pd.DataFrame]: Sequence of merged dataframes.
        """
        headers = ["&".join(df.columns.tolist()) for df in dfs]
        header_df = pd.DataFrame(headers, columns=["headers"])
        header_groups = header_df.groupby(by="headers")

        merged_dfs = []
        for _, df in header_groups:
            index = df.index.tolist()

            group = [dfs[i] for i in index]
            merged = pd.concat(group)
            merged.drop_duplicates(inplace=True)
            merged.reset_index(drop=True, inplace=True)
            merged_dfs.append(merged)

        return merged_dfs
