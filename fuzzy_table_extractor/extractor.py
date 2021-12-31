import os
from abc import abstractmethod
from dataclasses import dataclass
from datetime import datetime
from enum import Enum, auto
from pprint import pprint
from typing import List

import numpy as np
import pandas as pd

from ..Util.constants import MIN_HEADER_RATIO, VERBOSE
from ..Util.functions import (extract_datetime, format_sku, match_regex_list,
                              str_comparison)
from ..WordFile import WordFile
from .ExtractorsExceptions import ItemsNotFoundError, ApplicableItemNotFoundError


@dataclass
class HeaderInfo:
    NumeroED: str
    NumeroChamado: str
    Solicitante: str
    ObjED: str
    Elaborador: str
    VersaoED: str
    date: str


@dataclass
class TableMatch:
    search_term: str
    original_term: str
    score: str


@dataclass
class ReportDf:
    name: str
    df: pd.DataFrame


class FieldOrientation(Enum):
    ROW = auto()
    COLUMN = auto()


def sample_validation_function(header: List[str]) -> bool:
    return True


class Extractor:
    REPORT_PREFIX = '*'

    def __init__(self,  word_file: WordFile = None) -> None:
        self.word_file = word_file

    def extract_images_table(self) -> ReportDf:
        """Get a dataframe with document images data

        Returns:
            ReportDf: images dataframe
        """
        images = self.word_file.get_images_data()
        images = ["data:image/jpeg;base64, " + data for data in images]

        df = pd.DataFrame(images, columns=['images'])
        df['NumeroED'] = self.word_file.file_name

        df = df[['NumeroED', 'images']]

        return df

    def extract_headers_table(self) -> pd.DataFrame:
        """Default extraction of headers table for a report

        Returns:
            pd.DataFrame: report's header information
        """
        date = self.extract_single_field('Data',
                                         orientation=FieldOrientation.COLUMN,
                                         regex=[
                                             '\d+/\d+/\d+',
                                             '\d+-\d+-\d+',
                                         ])

        try:
            date = extract_datetime(date)
            date = datetime.strftime(date, '%d/%m/%Y')
        except ValueError:
            date = ''

        NumeroChamado = self.extract_single_field('Documento de Solicitação',
                                                  orientation=FieldOrientation.ROW)
        try:
            aux = NumeroChamado.split(':')[-1]
            if aux != '':
                NumeroChamado = aux
        except:
            pass

        Solicitante = self.extract_single_field('Solicitante',
                                                orientation=FieldOrientation.ROW)

        ObjED = self.extract_single_field('Objetivo',
                                          orientation=FieldOrientation.ROW)

        # TODO this won't work because this information is not in a table
        Elaborador = self.extract_single_field('elaborador',
                                               orientation=FieldOrientation.ROW)

        VersaoED = self.extract_single_field('VERSÃO',
                                             orientation=FieldOrientation.COLUMN)

        header_data = HeaderInfo(NumeroED=self.word_file.file_name,
                                 NumeroChamado=NumeroChamado,
                                 Solicitante=Solicitante,
                                 ObjED=ObjED,
                                 Elaborador=Elaborador,
                                 VersaoED=VersaoED,
                                 date=date)

        df = pd.DataFrame([header_data.__dict__])

        if VERBOSE:
            print('\n\nHEADERS')
            print(df)

        return df

    def extract_items_table(self) -> pd.DataFrame:
        """Default extraction of items table for a report

        Returns:
            pd.DataFrame: Report's items data
        """
        def validate_function(header: List[str]) -> bool:
            aux = ['lot' in x.lower() for x in header]

            return any(aux)

        try:
            df = self.extract_closest_table(search_headers=['Código', 'Descrição', 'Lote'],
                                            tables=self.word_file.dfs,
                                            validation_funtion=validate_function)

            if df is None:
                return None

            df.drop(columns='Descrição', inplace=True)
            df.rename(columns={'Código': 'Item',
                               'Lote': 'Lote'},
                      inplace=True)

            df.loc[:, 'NumeroED'] = self.word_file.file_name

            df = df[['NumeroED', 'Item', 'Lote']]
            df['Item'] = df['Item'].apply(format_sku)

            if VERBOSE:
                print('\n\nITENS')
                print(df)

            return df
        except:
            raise ItemsNotFoundError

    def extract_applicable_items_table(self) -> pd.DataFrame:
        """Default extraction of applicable items table for a report

        Returns:
            pd.DataFrame: applicable items data
        """
        def validate_function(header: List[str]) -> bool:
            aux = ['lot' in x.lower() for x in header]

            return not any(aux)

        try:
            df = self.extract_closest_table(search_headers=['Código', 'Descrição'],
                                            tables=self.word_file.dfs,
                                            validation_funtion=validate_function)

            if df is None:
                return None

            df.drop(columns='Descrição', inplace=True)
            df.rename(columns={'Código': 'Item'},
                      inplace=True)

            df.loc[:, 'NumeroED'] = self.word_file.file_name

            df = df[['NumeroED', 'Item']]
            df['Item'] = df['Item'].apply(format_sku)

            if VERBOSE:
                print('\n\nITENS APLICAVEIS')
                print(df)

            return df
        except:
            raise ApplicableItemNotFoundError

    @abstractmethod
    def extract_results_table(self) -> pd.DataFrame:
        """Extract the results table for a report
        For each report type, this function should be implemented

        Returns:
            pd.DataFrame: [description]
        """
        pass

    def get_all_dfs(self) -> List[ReportDf]:
        """Get a list of all dataframes for report, with table name of each

        The following tables are extracted:
            - headers
            - items
            - results
            - images
            - applicable items

        Returns:
            List[ReportDf]: list of dataframes in report
        """
        self.headers = self.extract_headers_table()
        self.items = self.extract_items_table()
        self.applicable_items = self.extract_applicable_items_table()
        self.results = self.extract_results_table()
        self.images = self.extract_images_table()

        header = ReportDf(name=f'{self.REPORT_PREFIX}_info',
                          df=self.headers)

        items = ReportDf(name=f'{self.REPORT_PREFIX}_items',
                         df=self.items)

        applicable_items = ReportDf(name=f'{self.REPORT_PREFIX}_applicable_items',
                                    df=self.applicable_items)

        results = ReportDf(name=f'{self.REPORT_PREFIX}_results',
                           df=self.results)

        images = ReportDf(name=f'{self.REPORT_PREFIX}_images',
                          df=self.images)

        return [header, items, applicable_items, results, images]

    def extract_closest_table(self,
                              search_headers: List[str],
                              tables: List[pd.DataFrame],
                              validation_funtion=sample_validation_function) -> pd.DataFrame:
        """Extract the table in document that has the closest header to search_headers

        Args:
            search_headers (List[str]): list of itens to search in header
            tables (List[pd.DataFrame]): list of dataframes in document

        Returns:
            pd.DataFrame: best match
        """
        ratios = []
        for df in tables:
            if validation_funtion(df.columns.to_list()):
                ratio = self.headers_proximity_ratio(document_headers=df.columns.to_list(),
                                                     search_headers=search_headers)

                ratios.append(ratio)
            else:
                ratios.append(0)

        if len(ratios) == 0:
            return None

        best_ratio = np.max(ratios)
        if best_ratio < MIN_HEADER_RATIO:
            return None

        best_match = tables[np.argmax(ratios)]

        association = self.headers_association(
            best_match.columns.to_list(), search_headers)

        if VERBOSE:
            pprint(association)
            print(best_ratio)

        original = [x.original_term for x in association]
        df = best_match[original]

        rename_columns = {x.original_term: x.search_term for x in association}
        df.rename(columns=rename_columns, inplace=True)

        return df

    def extract_single_field(self,
                             field: str,
                             orientation: FieldOrientation,
                             regex: List[str] = [''],
                             title_regex: List[str] = [''],
                             return_multiple: bool = False) -> str:
        """Extract single field of a word document based on a input string.
        The data will be extracted from tables in document

        Args:
            field (str): search field
            orientation (FieldOrientation): orientation to search the content of field
            regex (List[str], optional): list of regex to apply to content. To be a valid content there must be at least one match of regex in list. Defaults to [''].
            title_regex (List[str], optional): list of regex to apply to title. To be a valid title there must be at least one match of regex in list. Defaults to [''].
            return_multiple (bool, optional): if True, will return all matches that has the same proximity ratio. Defaults to False.

        Returns:
            str: best match
        """
        df = self.word_file.dictionary

        df = df[df['orientation'] == orientation.name.lower()]
        df = df[df['content'].apply(lambda x: match_regex_list(x, regex))]
        df = df[df['title'].apply(lambda x: match_regex_list(x, title_regex))]

        if df.empty:
            return ''

        df['ratio'] = df['title'].apply(lambda x: str_comparison(x, field))
        df.sort_values(by='ratio', inplace=True, ascending=False)

        try:
            if return_multiple:
                max_ratio = df['ratio'].max()
                values = df[df['ratio'] == max_ratio]['content'].to_list()
                best_match = ', '.join(values)
            else:
                best_match = df['content'].values[0]
        except IndexError:
            best_match = ''

        return best_match

    @staticmethod
    def headers_proximity_ratio(document_headers: List[str], search_headers: List[str]) -> int:
        """Calculates a proximity ratio of two headers

        Args:
            document_headers (List[str]): headers in document
            search_headers (List[str]): search headers

        Returns:
            int: proximity ratio
        """
        matches = Extractor.headers_association(document_headers,
                                                search_headers)

        if len(matches) == 0:
            return 0

        scores = [x.score for x in matches]

        return min(scores)

    @staticmethod
    def headers_association(document_headers: List[str], search_headers: List[str]) -> List[TableMatch]:
        # TODO I think this can be improved
        """Determine the best association of two headers

        Args:
            document_headers (List[str]): headers in document
            search_headers (List[str]): search headers

        Returns:
            List[TableMatch]: list of table headers matches
        """
        if len(search_headers) > len(document_headers):
            return []

        matches = []

        for s_header in search_headers:
            scores = [str_comparison(x, s_header) for x in document_headers]

            max_index = np.argmax(scores)
            max_score = np.max(scores)

            entry = TableMatch(search_term=s_header,
                               original_term=document_headers[max_index],
                               score=max_score)

            matches.append(entry)

            document_headers.pop(max_index)

        return matches

    @staticmethod
    def get_columns_fuzzy(df: pd.DataFrame, columns: List[str], threshold=0) -> pd.DataFrame:
        """Get columns that hat the closest match with supplied columns names
        The columns will be renamed to match the closest column name

        Args:
            df (pd.DataFrame): dataframe to search columns
            columns (List[str]): columns to search
            threshold (int, optional): minimum score to consider a match. Defaults to 0.

        Returns:
            List[str]: columns that match
        """
        association = Extractor.headers_association(df.columns.to_list(),
                                                    columns)

        association = [x for x in association if x.score > threshold]

        original = [x.original_term for x in association]
        df = df[original]

        rename_dict = {x.original_term: x.search_term for x in association}
        df.rename(columns=rename_dict, inplace=True)

        return df
