# coding: utf-8
from __future__ import annotations, unicode_literals
from typing import Any, cast
import scriptforge as SF

from ooodev.utils.session import Session


def _set_share_paths() -> None:
    # necessary hack for LO/python to import form "My Macros"
    # this method only runs once per session
    Session.register_path(Session.PathEnum.SHARE_USER_PYTHON)


_set_share_paths()

# Your code follows here
import pyglobal


def start_dialog() -> None:
    # casting is only for typing (autocomplete) support during design time. Cast is ignored at runtime.
    ui = cast(SF.SFScriptForge.SF_UI, SF.CreateScriptService("UI"))
    o_dialog = cast(SF.SFDialogs.SF_Dialog, SF.CreateScriptService("Dialog", ui.ActiveWindow, "Standard", "Dialog1"))

    lab_count = o_dialog.Controls("labCount")
    lab_count.Caption = str(pyglobal.G_COUNT)
    o_dialog.Execute()
    o_dialog.Terminate()


def updateLabel(event: Any, amt: int) -> None:
    btn_ctl = SF.CreateScriptService("DialogEvent", event)
    lab_count = btn_ctl.Parent.Controls("labCount")
    pyglobal.G_COUNT += amt
    lab_count.Caption = str(pyglobal.G_COUNT)
