# coding: utf-8
from enum import IntEnum, Enum


class CompareEnum(IntEnum):
    """Compare Enumeration"""

    AFTER = 1
    BEFORE = -1
    EQUAL = 0


class CompareEnum(IntEnum):
    """Compare Enumeration"""

    AFTER = 1
    BEFORE = -1
    EQUAL = 0


class AppTypeEnum(str, Enum):
    """LibreOffice App Type"""

    NONE = "NONE"
    UNKNOWN = "UNKNOWN"
    WRITER = "WRITER"
    CALC = "CALC"
    IMPRESS = "IMPRESS"
    DRAW = "DRAW"
    MATH = "MATH"
    BASE = "BASE"
