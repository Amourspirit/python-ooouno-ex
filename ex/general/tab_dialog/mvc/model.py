# coding: utf-8
from typing import List
from .interface import IModelMuiltiSyntax
from .enums import SyntaxEnum, VoidEnum


class MultiSyntaxModel(IModelMuiltiSyntax):

    def get_syntax_list_data(self) -> List[str]:
        """Get Main List Data"""
        return sorted([e.name for e in sorted(SyntaxEnum)])

    def get_void_list_data(self) -> List[str]:
        """Get Main List Data"""
        return sorted([e.name for e in sorted(VoidEnum)])
