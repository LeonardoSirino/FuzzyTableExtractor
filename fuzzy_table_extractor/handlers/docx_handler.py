import logging
import os
import re
from collections.abc import Sequence
from functools import cached_property
from pathlib import Path
from typing import List

import numpy as np
import pandas as pd
import win32com.client as win32
from docx.api import Document
from docx.table import Table

from ..util import table_to_dataframe
from .base_handler import BaseHandler

TITLE_NUM_IDS = [5, 7, 11]


class DocxHandler(BaseHandler):
    """The DocxHandler is handler for Microsoft Word Documents.
    It is suposed to use the newer .docx format for Word Documents, but if a .doc file is provided,
    it will be converted to .docx format.

    """

    def __init__(self, file_path: Path, temp_folder: Path = Path("temp")) -> None:
        """Creates a new DocxHandler instance.

        Args:
            file_path (Path): path to file to be handled
            temp_folder (str, optional): path to temporary folder where docx files will be created when the supplied file has .doc extension. Defaults to 'temp'.
        """
        super().__init__(file_path)
        self._file_path = file_path
        self._temp_folder = temp_folder

    @cached_property
    def words(self) -> Sequence[str]:
        document = self.document
        words = []

        # TODO check if there is another source of text that can be used for data extraction

        # Adding text from paragraphs
        paragraphs = document.paragraphs
        for item in paragraphs:
            aux = item.text
            aux = aux.split()
            words.extend(aux)

        # Adding text from tables
        for table in document.tables:
            for row in table.rows:
                data = [cell.text for cell in row.cells]
                for item in data:
                    words.extend(item.split())

        words = list(set(words))
        words = [word.lower() for word in words]
        words = [re.sub(r"[,]$|[.]$|[\(|\)|\:|\;|\'|\"]", "", word) for word in words]

        words = list(np.unique(words))

        return words

    @cached_property
    def dictionary(self) -> pd.DataFrame:
        logging.log(logging.INFO, "Getting dictionary from document")
        tables = self.docx_tables
        data = []
        for table in tables:
            n_cols = len(table.columns)

            for row in table.rows:
                cells = row.cells
                values = [cell.text for cell in cells]
                for k in range(n_cols - 1):
                    try:
                        title = values[k]
                        content = values[k + 1]
                        if title != content and title != "" and content != "":
                            data.append(
                                {
                                    "title": title,
                                    "content": content,
                                    "orientation": "row",
                                }
                            )
                    except IndexError:
                        pass

        for table in tables:
            n_rows = len(table.rows)

            for col in table.columns:
                cells = col.cells
                values = [cell.text for cell in cells]
                for k in range(n_rows - 1):
                    try:
                        title = values[k]
                        content = values[k + 1]
                        if title != content and title != "" and content != "":
                            data.append(
                                {
                                    "title": title,
                                    "content": content,
                                    "orientation": "column",
                                }
                            )
                    except IndexError:
                        pass

        df = pd.DataFrame(data)
        df.drop_duplicates(inplace=True)

        return df

    @cached_property
    def tables(self) -> Sequence[pd.DataFrame]:
        logging.log(logging.INFO, "Getting tables from document")
        tables = self.docx_tables

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

    @cached_property
    def docx_tables(self) -> List[Table]:
        """List of tables in docx document"""
        document = self.document
        tables = document.tables

        new_tables = tables
        count = 0

        # getting all inner tables
        while len(new_tables) > 0:
            count += 1
            aux = []
            for table in new_tables:
                for row in table.rows:
                    for cell in row.cells:
                        aux.extend(cell.tables)

            tables.extend(aux)
            new_tables = aux

            if count > 10:
                break

        return tables

    @cached_property
    def document(self) -> Document:
        """Open document and creates a docx file if necessary

        Returns:
            Document: word document object
        """
        folder, file_name = os.path.split(self._file_path)
        self.file_name = self._file_path.stem
        file_extension = self._file_path.suffix[1:]

        if file_extension == "doc":
            docx_file_name = f"x_{file_name}x"
            docx_file_path = self._temp_folder / docx_file_name

            if os.path.exists(docx_file_path):
                pass
            else:
                self._temp_folder.mkdir(exist_ok=True, parents=True)

                # Create a docx file if it yet does not exist in folder
                logging.info("Creating the docx file in the aux folder")

                # NOTE
                # If errors are found, do this
                # clear contents of C:\Users\<username>\AppData\Local\Temp\gen_py

                # Opening MS Word
                word = win32.gencache.EnsureDispatch("Word.Application")

                # NOTE this function does not accept the path as an object
                doc = word.Documents.Open(str(self._file_path.resolve()))
                doc.Activate()

                # Save and Close
                word.ActiveDocument.SaveAs(
                    str(docx_file_path.resolve()),
                    FileFormat=win32.constants.wdFormatXMLDocument,
                )
                doc.Close(False)

            self._file_path = docx_file_path

        document = Document(str(self._file_path.resolve()))
        return document

    def _merge_dfs(self, dfs: List[pd.DataFrame]) -> List[pd.DataFrame]:
        """Merge dataframes that has the same header and drop duplicated lines

        Args:
            dfs (List[pd.DataFrame]): list of dataframes from doc extraction

        Returns:
            List[pd.DataFrame]: list of merged dataframes
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
