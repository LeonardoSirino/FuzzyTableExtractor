import functools

import pandas as pd
from docx.table import Table
from fuzzywuzzy import fuzz
from unidecode import unidecode


def table_to_dataframe(table: Table) -> pd.DataFrame:
    data = []
    headers = []
    for i, row in enumerate(table.rows):
        row = [cell.text for cell in row.cells]

        if i == 0:
            headers = row
        else:
            data.append(row)

    # TODO sometimes the lenght of headers is different from the lenght of data
    # Check these cases and find a way to handle this situation
    df = pd.DataFrame(columns=headers, data=data)

    return df


@functools.cache
def str_comparison(text_a: str, text_b: str) -> int:
    """Get the proximity ratio of 2 strings

    Args:
        text_a (str): input text A
        text_b (str): input text B

    Returns:
        int: ration of proximity for 2 inputs [0 - 100]
    """

    if not isinstance(text_a, str) or not isinstance(text_b, str):
        return 0

    a = text_a.replace(" ", "").replace("\n", "")
    b = text_b.replace(" ", "").replace("\n", "")

    a = unidecode(a.lower())
    b = unidecode(b.lower())

    return fuzz.partial_ratio(a, b)
