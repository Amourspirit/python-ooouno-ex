# coding: utf-8
from typing import List, Union
from .interface import IViewMultiSyntax, IControllerMultiSyntax, IControllerSyntax
from .enums import TextEnum, TimeEnum, SyntaxEnum
from ...message_box.lib.msgbox import msgbox, MessageBoxType

class ControllerSyntax(IControllerSyntax):
    def __init__(self, controller: IControllerMultiSyntax, view: IViewMultiSyntax):
        self._controller = controller
        self._model = self._controller.model
        self._view = view
        self._dpv = False
        self._time = TimeEnum.NONE
        self._text = TextEnum.NONE
        self._syntax: Union[SyntaxEnum, None] = None


    def get_list_data(self) -> List[str]:
        return self._model.get_syntax_list_data()

    # region Handler Methods

    def write(self):
        if self._syntax:
            msg = f"Selected: {self._syntax.name}\nTime: {self.time.name}\nSuffix: {self.text.name}"
            msg += f"\nDPV: {self.dpv}"
            
            msgbox(
                message=msg,
                title="SYNTAX",
                boxtype=MessageBoxType.INFOBOX,
            )
        else:
            msgbox("Nothing selected.", title="SYNTAX", boxtype=MessageBoxType.WARNINGBOX)

    # endregion Handler Methods

    def is_syntax(self) -> bool:
        """
        Gets if Property ``syntax`` is None or ``SyntaxEnum``

        Returns:
            bool: ``False`` if ``syntax`` is ``None``; Otherwise, ``True``.
        """
        return not self._syntax is None

    # region Properties

    @property
    def dpv(self) -> bool:
        """Gets/Sets dpv state"""
        return self._dpv

    @dpv.setter
    def dpv(self, value: bool):
        self._dpv = value
        self._view.refresh()

    @property
    def time(self) -> TimeEnum:
        """Gets/sets time state"""
        return self._time

    @time.setter
    def time(self, value: TimeEnum):
        self._time = value
        self._view.refresh()

    @property
    def text(self) -> TextEnum:
        """Gets/sets prefix suffix state"""
        return self._text

    @text.setter
    def text(self, value: TextEnum):
        self._text = value
        self._view.refresh()

    @property
    def syntax(self) -> Union[SyntaxEnum, None]:
        return self._syntax

    @syntax.setter
    def syntax(self, value: Union[SyntaxEnum, None]):
        self._syntax = value
        self._view.refresh()

    # endregion Properties
