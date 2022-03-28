# coding: utf-8
from src.examples.message_box.msgbox import msgbox, MessageBoxButtonsEnum, MessageBoxResultsEnum,MessageBoxType

def msg_small(*args, **kwargs):
    msg = "A small message"
    result = msgbox(
        message=msg,
        title="Short",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.INFOBOX,
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
    result = msgbox(
        message=msg,
        buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        title="Long...",
        boxtype=MessageBoxType.QUERYBOX,
    )
    assert result == MessageBoxResultsEnum.YES or MessageBoxResultsEnum.NO
    print(result)

def msg_default_yes(*args, **kwargs):
    msg = "This dialog as button set to a defalt of yes."
    result = msgbox(
        message=msg,
        buttons=(
            MessageBoxButtonsEnum.BUTTONS_YES_NO.value
            | MessageBoxButtonsEnum.DEFAULT_BUTTON_YES.value
        ),
        title="Default",
        boxtype=MessageBoxType.MESSAGEBOX,
    )
    assert result == MessageBoxResultsEnum.YES or MessageBoxResultsEnum.NO
    print(result)

def msg_error(*args, **kwargs):
    msg = "Looks like an error!"
    result = msgbox(
        message=msg,
        title="\U0001F6D1 Opps",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.ERRORBOX,
    )
    assert result == MessageBoxResultsEnum.OK
    print(result)

def msg_warning(*args, **kwargs):
    msg = "Looks like a Warning!"
    result = msgbox(
        message=msg,
        title="\u26A0 \U0001F440",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.WARNINGBOX,
    )
    assert result == MessageBoxResultsEnum.OK
    print(result)
