# coding: utf-8
from .msgbox import msgbox, MessageBoxButtonsEnum, MessageBoxResultsEnum,MessageBoxType

def main():
    msg = "A small message"
    result = msgbox(
        message=msg,
        title="Short",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.INFOBOX,
    )
    assert result == MessageBoxResultsEnum.OK
    print(result)
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

    msg = "Looks like an error!"
    result = msgbox(
        message=msg,
        title="\U0001F6D1 Opps",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.ERRORBOX,
    )
    assert result == MessageBoxResultsEnum.OK
    print(result)

    msg = "Looks like a Warning!"
    result = msgbox(
        message=msg,
        title="\u26A0 \U0001F440",
        buttons=MessageBoxButtonsEnum.BUTTONS_OK,
        boxtype=MessageBoxType.WARNINGBOX,
    )
    assert result == MessageBoxResultsEnum.OK
    print(result)