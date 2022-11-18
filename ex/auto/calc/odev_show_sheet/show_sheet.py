import uno

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr

from com.sun.star.util import XProtectable


class ShowSheet:
    def __init__(self, read_only: bool, input_fnm: PathOrStr, out_fnm: PathOrStr, visible: bool) -> None:
        self._readonly = read_only
        _ = FileIO.is_exist_file(input_fnm, True)
        self._input_fnm = FileIO.get_absolute_path(input_fnm)
        if out_fnm:
            outf = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(outf)
            self._out_fnm = outf
        else:
            self._out_fnm = ""
        self._visible = visible

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = Calc.open_doc(fnm=self._input_fnm, loader=loader)

            # doc = Lo.open_readonly_doc(fnm=self._input_fnm, loader=loader)
            # doc = Calc.get_ss_doc(doc)

            if self._visible:
                GUI.set_visible(is_visible=True, odoc=doc)


            Calc.goto_cell(cell_name="A1", doc=doc)
            sheet_names = Calc.get_sheet_names(doc=doc)
            print(f"Names of Sheets ({len(sheet_names)}):")
            for name in sheet_names:
                print(f"  {name}")

            sheet = Calc.get_sheet(doc=doc, index=0)
            Calc.set_active_sheet(doc=doc, sheet=sheet)
            pro = Lo.qi(XProtectable, sheet, True)
            pro.protect("foobar")
            print(f"Is protected: {pro.isProtected()}")

            Lo.delay(2000)
            # query the user for the password
            pwd = GUI.get_password("Password", "Enter sheet Password")
            if pwd == "foobar":
                pro.unprotect(pwd)
                MsgBox.msgbox("Password is Correct", "Password", boxtype=MessageBoxType.INFOBOX)
            else:
                MsgBox.msgbox("Password is incorrect", "Password", boxtype=MessageBoxType.ERRORBOX)

            if self._out_fnm:
                Lo.save_doc(doc=doc, fnm=self._out_fnm)

            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                Lo.close_doc(doc=doc, deliver_ownership=True)
                Lo.close_office()
            else:
                print("Keeping document open")

        except Exception:
            Lo.close_office()
            raise
