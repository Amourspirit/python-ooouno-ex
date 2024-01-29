from __future__ import annotations
import uno

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import CalcDoc
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr

from com.sun.star.util import XProtectable


class ShowSheet:
    def __init__(
        self, read_only: bool, input_fnm: PathOrStr, out_fnm: PathOrStr, visible: bool
    ) -> None:
        self._readonly = read_only
        _ = FileIO.is_exist_file(input_fnm, True)
        self._input_fnm = FileIO.get_absolute_path(input_fnm)
        if out_fnm:
            out_file = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(out_file)
            self._out_fnm = out_file
        else:
            self._out_fnm = ""
        self._visible = visible

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = CalcDoc.open_doc(
                fnm=self._input_fnm, loader=loader, visible=self._visible
            )

            sheet = doc.get_active_sheet()

            sheet["A1"].goto()
            sheet_names = doc.get_sheet_names()
            print(f"Names of Sheets ({len(sheet_names)}):")
            for name in sheet_names:
                print(f"  {name}")

            sheet.set_active()
            pro = sheet.qi(XProtectable, True)
            pro.protect("foobar")
            print(f"Is protected: {pro.isProtected()}")

            Lo.delay(2000)
            # query the user for the password
            pwd = GUI.get_password("Password", "Enter sheet Password")
            if pwd == "foobar":
                pro.unprotect(pwd)
                MsgBox.msgbox(
                    "Password is Correct", "Password", boxtype=MessageBoxType.INFOBOX
                )
            else:
                MsgBox.msgbox(
                    "Password is incorrect", "Password", boxtype=MessageBoxType.ERRORBOX
                )

            if self._out_fnm:
                doc.save_doc(fnm=self._out_fnm)

            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.close_doc()
                Lo.close_office()
            else:
                print("Keeping document open")

        except Exception:
            Lo.close_office()
            raise
