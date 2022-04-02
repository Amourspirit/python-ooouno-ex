# coding: utf-8
from enum import IntEnum, auto
from ....utils import enum_util as eutil

class TimeEnum(IntEnum):
    """Time Enum Values"""

    NONE: int = 1
    """No time Selection"""
    PAST: int = 8
    """Time Past"""
    FUTURE: int = 9
    """Time_future"""

    def __str__(self):
        return self._name_


class TextEnum(IntEnum):
    """Text Enum Values"""

    NONE: int = 0
    """No text selection"""
    PREFIX: int = 1
    """Text Prefix"""
    SUFFIX: int = 2
    """Text_suffix"""

    def __str__(self):
        return self._name_


class TabEnum(IntEnum):
    """Tabs Enum"""

    SYNTAX: int = 1
    """Syntax tab"""
    VOID: int = 2
    """Void tab"""

class SyntaxEnum(IntEnum):
    ARRAGNEMENT = 1
    SYNTAX = 2
    ORDER = 3
    PATTERN = 4
    STRUCTURE =5
    PLAN = 6
    METHOD = 7
    STRATEGY = 8

    def __str__(self):
        return self._name_

# constructor of Enum can not be changed until after the enum is intially constructed.
# By setting __new__ it is possible to construct from name.
setattr(SyntaxEnum, '__new__', eutil.enum_class_new_flex)

class VoidEnum(IntEnum):
    VOID = auto()
    VACANT = auto()
    VACUUM = auto()
    UNOCCUPIED = auto()
    BLANK = auto()
    ABANDONED = auto()
    BARE = auto()
    BARREN = auto()

setattr(VoidEnum, '__new__', eutil.enum_class_new_flex)