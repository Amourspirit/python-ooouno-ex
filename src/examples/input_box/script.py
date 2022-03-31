# coding: utf-8
from src.examples.message_box.msgbox import msgbox, MessageBoxButtonsEnum, MessageBoxResultsEnum,MessageBoxType
from src.examples.input_box.inputbox import InputBox


def show_msg_value(msg: str) -> None:
    msgbox(
        message=msg,
        title="Input box Result",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.INFOBOX,
    )

def show_msg_no_value() -> None:
    msgbox(
        message="No input",
        title="Input box Result",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.WARNINGBOX,
    )

def input_box(*args, **kwargs):
    box = InputBox("", title="Input")
    result = box.show()
    if len(result) == 0:
        show_msg_no_value()
    else:
        show_msg_value(result)

