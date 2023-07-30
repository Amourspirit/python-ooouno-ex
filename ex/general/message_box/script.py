from ooodev.macro.macro_loader import MacroLoader
from ooodev.dialog.msgbox import MsgBox,  MessageBoxButtonsEnum, MessageBoxResultsEnum, MessageBoxType


def msg_small(*args, **kwargs):
    msg = "A small message"
    with MacroLoader():
        result = MsgBox.msgbox(
            msg=msg,
            title="Short",
            boxtype=MessageBoxType.INFOBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        )
        assert result == MessageBoxResultsEnum.OK
        print(result)


def msg_long(*args, **kwargs):
    msg = (
        "A very long message A very long message A very long message A very long message "
        "A very long message A very long message A very long message A very long message "
        "A very long message A very long message"
        "\n\n"
        "Do you agree ?"
    )
    with MacroLoader():
        result = MsgBox.msgbox(
            msg=msg,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            boxtype=MessageBoxType.QUERYBOX,
            title="Long...",
        )
        assert result == MessageBoxResultsEnum.YES or MessageBoxResultsEnum.NO
        print(result)


def msg_default_yes(*args, **kwargs):
    msg = "This dialog as button set to a default of yes."
    with MacroLoader():
        result = MsgBox.msgbox(
            msg=msg,
            buttons=(MessageBoxButtonsEnum.BUTTONS_YES_NO.value | MessageBoxButtonsEnum.DEFAULT_BUTTON_YES.value),
            title="Default",
            boxtype=MessageBoxType.MESSAGEBOX,
        )
        assert result == MessageBoxResultsEnum.YES or MessageBoxResultsEnum.NO
        print(result)


def msg_error(*args, **kwargs):
    msg = "Looks like an error!"
    with MacroLoader():
        result = MsgBox.msgbox(
            msg=msg,
            title="\U0001F6D1 Opps",
            boxtype=MessageBoxType.ERRORBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        )
        assert result == MessageBoxResultsEnum.OK
        print(result)


def msg_warning(*args, **kwargs):
    msg = "Looks like a Warning!"
    with MacroLoader():
        result = MsgBox.msgbox(
            msg=msg,
            title="\u26A0 \U0001F440",
            boxtype=MessageBoxType.WARNINGBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        )
        assert result == MessageBoxResultsEnum.OK
        print(result)
