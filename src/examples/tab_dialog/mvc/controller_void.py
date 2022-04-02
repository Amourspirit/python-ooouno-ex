# coding: utf-8
from typing import List, Union
from .interface import IViewMultiSyntax, IControllerMultiSyntax, IControllerVoid
from .enums import VoidEnum
from ...message_box.msgbox import msgbox, MessageBoxType


class ControllerVoid(IControllerVoid):
    def __init__(self, controller: IControllerMultiSyntax, view: IViewMultiSyntax):
        self._controller = controller
        self._model = self._controller.model
        self._view = view
        self._callbacks = None
        self._void: Union[VoidEnum, None] = None

    def write(self):
        if self._void:
            msgbox(
                message=f"Selected Void Value of {self._void.name}",
                title="VOID",
                boxtype=MessageBoxType.INFOBOX,
            )
        else:
            msgbox("Nothing selected.", title="VOID", boxtype=MessageBoxType.WARNINGBOX)

    def get_list_data(self) -> List[str]:
        return self._model.get_void_list_data()

    @property
    def void(self) -> Union[VoidEnum, None]:
        return self._void

    @void.setter
    def void(self, value: Union[VoidEnum, None]):
        self._void = value
        self._view.refresh()
