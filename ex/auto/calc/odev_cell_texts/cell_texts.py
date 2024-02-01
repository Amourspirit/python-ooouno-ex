from __future__ import annotations

import uno
from com.sun.star.text import XText
from com.sun.star.text import XSentenceCursor
from com.sun.star.text import XParagraphCursor

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.format.calc.direct.cell.borders import Borders, Padding
from ooodev.format.calc.direct.cell.font import Font
from ooodev.calc import CalcDoc
from ooodev.write import Write
from ooodev.units.unit_mm import UnitMM
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class CellTexts:
    def __init__(self, out_fnm: PathOrStr, **kwargs) -> None:
        if out_fnm:
            out_f = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(out_f)
            self._out_fnm = out_f
        else:
            self._out_fnm = ""

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = CalcDoc.create_doc(loader=loader, visible=True)

            sheet = doc.sheets[0]

            rng = sheet.get_range(range_name="A2:D7")
            _ = rng.highlight(headline="Cells and Cell Ranges")

            # Insert two text paragraphs and a hyperlink into the cell

            cell = sheet["B4"]
            x_text = cell.qi(XText, True)
            cursor = x_text.createTextCursor()
            Write.append_para(cursor=cursor, text="Text in first line.")
            Write.append(cursor=cursor, text="And a ")
            Write.add_hyperlink(
                cursor=cursor,
                label="hyperlink",
                url_str="https://github.com/Amourspirit/python_ooo_dev_tools",
            )

            # beautify the cell
            font = Font(color=CommonColor.DARK_BLUE, size=18.0)
            bdr = Borders(padding=Padding(left=UnitMM(5)))
            cell.apply_styles(font, bdr)

            self._print_cell_text(x_text)

            cell.add_annotation(msg=f"This annotation is located at {cell.cell_obj}")

            if self._out_fnm:
                doc.save_doc(fnm=self._out_fnm)

            msg_result = doc.msgbox(
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

    def _print_cell_text(self, cell_text: XText) -> None:
        print(f'Cell Text: "{cell_text.getString()}"')

        cursor = cell_text.createTextCursor()
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
