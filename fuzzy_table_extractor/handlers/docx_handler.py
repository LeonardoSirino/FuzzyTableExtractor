from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from functools import cache, lru_cache
from io import StringIO
from pathlib import Path
from typing import List, Optional

import numpy as np
import pandas as pd
import win32com.client as win32
from docx.api import Document
from docx.oxml.numbering import CT_NumPr
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_Text
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree

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

    @property
    @lru_cache
    def words(self) -> List[str]:
        logging.log(logging.INFO, "Getting words from document")
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

    @property
    @lru_cache
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

    @property
    @lru_cache
    def tables(self) -> List[pd.DataFrame]:
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

    @property
    @lru_cache
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

    @property
    @lru_cache
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
                # C:\Users\u122004\AppData\Local\Temp\gen_py

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


@dataclass
class _DocxTopic:
    level: int
    num_id: int
    text: str

    def __post_init__(self) -> None:
        text = self.text

        patterns_to_remove = [
            r"\_{2,}",
            r"\;$",
        ]
        for pattern in patterns_to_remove:
            text = re.sub(pattern, "", text)

        text = text.strip()

        self.text = text

    def __repr__(self) -> str:
        text = f"{'   ' * self.level}{self.text}"

        return text


@dataclass
class _DocxCheckbox:
    name: str
    state: bool

    def __post_init__(self) -> None:
        name = self.name

        patterns_to_remove = [
            r"\_{2,}",
        ]
        for pattern in patterns_to_remove:
            name = re.sub(pattern, "", name)

        name = name.strip()

        self.name = name


@dataclass
class _DocxHyperlink:
    text: str
    level: int
    xpath: str

    def __repr__(self) -> str:
        text = f"{'  ' * self.level}{self.text}"
        return text


class _DocxHyperlinksManager:
    def __init__(self, document_content: str) -> None:
        self._hyperlinks: List[_DocxHyperlink] = []
        self._parse(document_content)

    def _parse(self, content: str) -> None:
        io_content = StringIO(content)
        self.tree = etree.parse(io_content, etree.XMLParser(remove_blank_text=True))

        root = self.tree.getroot()
        numbers = root.xpath("//w:hyperlink/w:r[1]/w:t", namespaces=root.nsmap)

        titles = root.xpath("//w:hyperlink/w:r[3]/w:t", namespaces=root.nsmap)

        for number, title in zip(numbers, titles):
            xpath = self.tree.getpath(title)
            level = len([x for x in number.text if x == "."]) - 1
            hyperlink = _DocxHyperlink(title.text.strip(), level, xpath)

            self._hyperlinks.append(hyperlink)

    def has_items(self) -> bool:
        return len(self._hyperlinks) > 0

    def get_paragraph_level(self, par: Paragraph) -> int:
        """Checks if the supplied paragraph matches the next section name, returns the level of the section name if it does

        Args:
            par (Paragraph): paragraph to be checked

        Raises:
            ValueError: When the paragraph is not a section name or there is no more section names in the document

        Returns:
            int: the level of the section name
        """
        try:
            next_link = self._hyperlinks[0]
        except IndexError:
            raise ValueError("There are no more hyperlinks in the document")

        # TODO match the item xpath, not its content
        pattern = fr"^{next_link.text}"
        if re.search(pattern, par.text.strip(), re.IGNORECASE):
            logging.debug(f"Found hyperlink: {next_link}")
            level = next_link.level
            self._hyperlinks.pop(0)
            return level
        else:
            raise ValueError(
                f"Paragraph text '{par.text}' does not match next hyperlink '{next_link.text}'"
            )

    def __str__(self) -> str:
        if not self._hyperlinks:
            return "No hyperlinks found"

        blocks = []
        for hyperlink in self._hyperlinks:
            text = f"{'    ' * hyperlink.level}{hyperlink.text}"
            blocks.append(text)

        return "\n".join(blocks)


class _DocxNode:
    """DocxNode is used to create a tree like structure of a Microsoft Word Document."""

    def __init__(self, content: str, num_id: int, level: int):
        # TODO find a better way to handle line breaks in content
        self.content = content.strip().replace("\n", "|")
        self.num_id = num_id
        self.level = level

        self.paragraphs: List[str] = []
        self.tables: List[pd.DataFrame] = []
        self.topics: List[_DocxTopic] = []
        self.checkboxes: List[_DocxCheckbox] = []

        self.nodes: List[_DocxNode] = []

        self._last_node: _DocxNode = self

    def add_paragraph(self, paragraph: str):
        """Add the given paragraph to the last node added to the tree.

        Args:
            paragraph (str): paragraph content
        """
        self._last_node.paragraphs.append(paragraph)

    def add_topic(self, topic: _DocxTopic):
        """Add the given topic to the last node added to the tree.

        Args:
            topic (DocxTopic): topic to be added
        """
        self._last_node.topics.append(topic)

    def add_table(self, table: pd.DataFrame):
        """Add the given table to the last node added to the tree.

        Args:
            table (pd.DataFrame): table content as a pandas DataFrame
        """
        self._last_node.tables.append(table)

    def add_checkbox(self, checkbox: _DocxCheckbox):
        """Add the given checkbox to the last node added to the tree.

        Args:
            checkbox (DocxCheckbox): checkbox data
        """
        self._last_node.checkboxes.append(checkbox)

    def add_node(self, node: _DocxNode):
        """Finds the last node in tree that is one level below the current node and adds the given node to it."""
        parent_section = self
        while True:
            if node.level == parent_section.level + 1:
                break

            parent_section = parent_section.nodes[-1]

        parent_section.nodes.append(node)
        self._last_node = node

    def __repr__(self):
        children = ""
        for node in self.nodes:
            lines = str(node).split("\n")
            for line in lines:
                children += f"\t{line}\n"

        metrics = (
            f"{len(self.paragraphs)}P | {len(self.topics)}To | {len(self.tables)}Ta"
        )
        base = f"{self.level} - {self.content} ({metrics})"

        return f"{base}\n{children}"

    def get_topics(self, recursive: bool = False) -> List[_DocxTopic]:
        if not recursive:
            return self.topics

        topics = self.topics
        for node in self.nodes:
            topics.extend(node.get_topics(recursive))

        return topics

    def get_checkboxes(self, recursive: bool = False) -> List[_DocxCheckbox]:
        if not recursive:
            return self.checkboxes

        checkboxes = self.checkboxes
        for node in self.nodes:
            checkboxes.extend(node.get_checkboxes(recursive))

        return checkboxes

    def get_paragraphs(self, recursive: bool = False) -> List[str]:
        if not recursive:
            return self.paragraphs

        paragraphs = self.paragraphs
        for node in self.nodes:
            paragraphs.extend(node.get_paragraphs(recursive))

        return paragraphs

    def get_nodes(self, recursive: bool = False) -> List[_DocxNode]:
        if not recursive:
            return self.nodes

        nodes = self.nodes
        for node in self.nodes:
            nodes.extend(node.get_nodes(recursive))

        return nodes

    def get_tables(self, recursive: bool = False) -> List[pd.DataFrame]:
        if not recursive:
            return self.tables

        tables = self.tables
        for node in self.nodes:
            tables.extend(node.get_tables(recursive))

        return tables


class TreeDocxHandler(DocxHandler):
    """The TreeDocxHandler is a specialization of the DocxHandler that creates a tree structure of the document.
    Each level on this tree is related to a section in the document.
    This handler should be used when there is a need to access information in a specific section of the document, otherwise the 
    DocxHandler should be used, as it is more efficient.
    """

    def __init__(self, file_path: Path) -> None:
        super().__init__(file_path)
        self.file_name = file_path.stem

    def _handle_checkboxes(self, paragraph: Paragraph, root: _DocxNode) -> None:
        split_candidates = [
            "/",
            " / ",
            "-",
            " - ",
            "\n",
        ]

        checkboxes = paragraph._element.xpath(".//w:checkBox/w:default")
        states = []
        for x in checkboxes:
            _, value = x.items()[0]
            states.append(value == "1")

        if len(states) == 0:
            return
        elif len(states) == 1:
            root.add_checkbox(_DocxCheckbox(paragraph.text, states[0]))
        else:
            for split in split_candidates:
                names = paragraph.text.split(split)
                if len(names) == len(states):
                    for name, state in zip(names, states):
                        root.add_checkbox(_DocxCheckbox(name, state))

                    return

    def _handle_paragraph(self, paragraph: Paragraph, root: _DocxNode) -> None:
        # TODO also remove some chars not wanted to constitute a paragraph
        if paragraph.text.strip() == "":
            return

        self._handle_checkboxes(paragraph, root)

        # print(paragraph.text)
        # text = 'parecer final / conclusion'
        # if text in paragraph.text.lower():
        #     print(paragraph.text)
        #     print(paragraph._p.pPr.numPr)
        #     quit()

        # trying to find paragraph level as a hyperlink
        try:
            level = self._hyperlinks_manager.get_paragraph_level(paragraph)
            section = _DocxNode(paragraph.text, 5, level)
            root.add_node(section)

            return
        except ValueError:
            pass

        # trying to find paragraph level by its pPr number and level
        try:
            number: CT_NumPr = paragraph._p.pPr.numPr
            level = number.ilvl.val
            num_id = number.numId.val

            if num_id in TITLE_NUM_IDS:
                section = _DocxNode(paragraph.text, num_id, level)
                root.add_node(section)
            else:
                topic = _DocxTopic(level, num_id, paragraph.text)
                root.add_topic(topic)

            return
        except AttributeError:
            pass

        root.add_paragraph(paragraph.text)

    def _handle_table(self, table: Table, root: _DocxNode) -> None:
        # TODO handle inner tables that may appear in document
        df = table_to_dataframe(table)
        root.add_table(df)

    @property
    @cache
    def root_node(self) -> _DocxNode:
        """The root node of the document tree."""
        self._hyperlinks_manager = _DocxHyperlinksManager(self.document._element.xml)

        doc = self.document
        root = _DocxNode("root", num_id=-1, level=-1)
        for child in doc.element.body.iterchildren():
            if isinstance(child, CT_P):
                par = Paragraph(child, self.document)
                self._handle_paragraph(par, root)
            elif isinstance(child, CT_Tbl):
                table = Table(child, self.document)
                self._handle_table(table, root)
            else:
                logging.info(f"Unhandled child type: {type(child)}")

        if self._hyperlinks_manager.has_items():
            logging.warning("Not all hyperlinks were handled")

        return root
