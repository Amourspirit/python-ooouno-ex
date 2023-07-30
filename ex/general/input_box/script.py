from ooodev.dialog.input import Input
from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum
from ooodev.macro.macro_loader import MacroLoader


def show_msg_value(msg: str) -> None:
    _ = MsgBox.msgbox(
        msg=msg, title="Input box Result", boxtype=MessageBoxType.INFOBOX, buttons=MessageBoxButtonsEnum.BUTTONS_OK
    )


def show_msg_no_value() -> None:
    _ = MsgBox.msgbox(
        msg="No Input",
        title="Input box Result",
        boxtype=MessageBoxType.WARNINGBOX,
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
    )


def input_box(*args, **kwargs):
    with MacroLoader():
        result = Input.get_input(title="Input", msg="")
        if len(result) == 0:
            show_msg_no_value()
        else:
            show_msg_value(result)
