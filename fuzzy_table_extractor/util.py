import re
from functools import lru_cache
from typing import List

from fuzzywuzzy import fuzz
from unidecode import unidecode


@lru_cache(maxsize=None)
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

    ratio = fuzz.partial_ratio(a, b)

    return ratio


def match_regex_list(text: str, patterns: List[str]) -> bool:
    """Checks if text matches any regex in list
    The match is performed in a case insensitive way

    Args:
        text (str): input text
        regex (List[str]): list of regex

    Returns:
        bool: match of any regex
    """
    if len(patterns) == 0:
        return True

    for pattern in patterns:
        res = re.finditer(pattern, text, re.IGNORECASE)

        try:
            next(res)
            return True
        except StopIteration:
            pass

    return False
