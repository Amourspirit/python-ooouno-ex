# coding: utf-8
from __future__ import annotations, unicode_literals
import sys
import uno
from typing import Any, cast
import getpass, os, os.path
import scriptforge as SF


class Session:
    # https://help.libreoffice.org/latest/lo/text/sbasic/python/python_import.html
    # https://help.libreoffice.org/latest/lo/text/sbasic/python/python_session.html
    @staticmethod
    def substitute(var_name):
        ctx = uno.getComponentContext()
        ps = ctx.getServiceManager().createInstanceWithContext("com.sun.star.util.PathSubstitution", ctx)
        return ps.getSubstituteVariableValue(var_name)

    @staticmethod
    def Share():
        inst = uno.fileUrlToSystemPath(Session.substitute("$(prog)"))
        return os.path.normpath(inst.replace("program", "Share"))

    @staticmethod
    def SharedScripts():
        # eg: C:\Program Files\LibreOffice\share\Scripts
        return "".join([Session.Share(), os.sep, "Scripts"])

    @staticmethod
    def SharedPythonScripts():
        # eg: C:\Program Files\LibreOffice\share\Scripts\python
        return "".join([Session.SharedScripts(), os.sep, "python"])

    @property  # alternative to '$(username)' variable
    def UserName(self):
        return getpass.getuser()

    @property
    def UserProfile(self):
        return uno.fileUrlToSystemPath(Session.substitute("$(user)"))

    @property
    def UserScripts(self):
        # eg: C:\Users\user\AppData\Roaming\LibreOffice\4\user\Scripts
        return "".join([self.UserProfile, os.sep, "Scripts"])

    @property
    def UserPythonScripts(self):
        # eg: C:\Users\user\AppData\Roaming\LibreOffice\4\user\Scripts\python
        return "".join([self.UserScripts, os.sep, "python"])


def _set_share_paths() -> None:
    # necessary hack for LO/python to import from "My Macros"
    # this method only runs once per session
    ss_info = Session()
    print("UserPythonScripts", ss_info.UserPythonScripts)
    if not ss_info.UserPythonScripts in sys.path:
        sys.path.insert(0, ss_info.UserPythonScripts)  # Add to search path


_set_share_paths()

# Your code follows here
import pyglobal


def start_dialog(*args, **kwargs) -> None:
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


def btnIncrement(event: Any) -> None:
    updateLabel(event, 1)


def btnDecrement(event: Any) -> None:
    updateLabel(event, -1)


g_exportedScripts = (start_dialog, btnIncrement, btnDecrement)
