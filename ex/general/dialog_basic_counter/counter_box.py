from __future__ import annotations
import scriptforge as SF
from typing import Any, cast

G_COUNT = 100


def start_dialog(*args, **kwargs) -> None:
    global G_COUNT
    # casting is only for typing (autocomplete) support during design time. Cast is ignored at runtime.
    ui = cast(SF.SFScriptForge.SF_UI, SF.CreateScriptService("UI"))
    o_dialog = cast(SF.SFDialogs.SF_Dialog, SF.CreateScriptService("Dialog", ui.ActiveWindow, "Standard", "Dialog1"))

    # uncomment line to have count reset to 100 each time dialog loads.
    # G_COUNT = 100

    lab_count = o_dialog.Controls("labCount")
    lab_count.Caption = str(G_COUNT)
    o_dialog.Execute()
    o_dialog.Terminate()


def updateLabel(event: Any, amt: int) -> None:
    global G_COUNT
    btn_ctl = SF.CreateScriptService("DialogEvent", event)
    lab_count = btn_ctl.Parent.Controls("labCount")
    G_COUNT += amt
    lab_count.Caption = str(G_COUNT)


def btnIncrement(event: Any) -> None:
    updateLabel(event, 1)


def btnDecrement(event: Any) -> None:
    updateLabel(event, -1)


g_exportedScripts = (start_dialog, btnIncrement, btnDecrement)
