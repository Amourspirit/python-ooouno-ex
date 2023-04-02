from __future__ import annotations

import uno
from com.sun.star.text import XText
from com.sun.star.table import XCell
from com.sun.star.text import XSentenceCursor
from com.sun.star.text import XParagraphCursor

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.format import Styler
from ooodev.format.calc.direct.cell.borders import Borders, Padding
from ooodev.format.calc.direct.cell.font import Font
from ooodev.office.calc import Calc
from ooodev.office.write import Write
from ooodev.units.unit_mm import UnitMM
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class CellTexts:
    def __init__(self, out_fnm: PathOrStr, **kwargs) -> None:
        if out_fnm:
            outf = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(outf)
            self._out_fnm = outf
        else:
            self._out_fnm = ""

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = Calc.create_doc(loader)

            GUI.set_visible(is_visible=True, odoc=doc)

            sheet = Calc.get_sheet(doc=doc, index=0)

            Calc.highlight_range(sheet=sheet, range_name="A2:C7", headline="Cells and Cell Ranges")

            xcell = Calc.get_cell(sheet=sheet, cell_name="B4")

            # Insert two text paragraphs and a hyperlink into the cell
            xtext = Lo.qi(XText, xcell, True)
            cursor = xtext.createTextCursor()
            Write.append_para(cursor=cursor, text="Text in first line.")
            Write.append(cursor=cursor, text="And a ")
            Write.add_hyperlink(
                cursor=cursor, label="hyperlink", url_str="https://github.com/Amourspirit/python_ooo_dev_tools"
            )

            # beautify the cell
            font = Font(color=CommonColor.DARK_BLUE, size=18.0)
            bdr = Borders(padding=Padding(left=UnitMM(5)))
            Styler.apply(xcell, font, bdr)

            self._print_cell_text(xcell)

            Calc.add_annotation(sheet=sheet, cell_name="B4", msg="This annotation is located at B4")

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

    def _print_cell_text(self, cell: XCell) -> None:
        txt = Lo.qi(XText, cell, True)
        print(f'Cell Text: "{txt.getString()}"')

        cursor = txt.createTextCursor()
        if cursor is None:
            print("Text cursor is null")
            return

        sent_cursor = Lo.qi(XSentenceCursor, cursor)
        if sent_cursor is None:
            print("Sentence cursor is null")

        para_cursor = Lo.qi(XParagraphCursor, cursor)
        if para_cursor is None:
            print("Paragraph cursor is null")
            return

        # go to start of text; no selection
        para_cursor.gotoStart(False)

        while True:
            #  select all of paragraph
            para_cursor.gotoEndOfParagraph(True)
            print(para_cursor.getString())
            if not para_cursor.gotoNextParagraph(False):
                break
