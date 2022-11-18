from __future__ import annotations

import uno
from com.sun.star.drawing import XDrawPagesSupplier
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.lang import XComponent
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc
from ooodev.office.draw import Draw
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.type_var import PathOrStr

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    # bug in LibreOffice 7.4
    Chart2 = None

from ooo.dyn.table.cell_hori_justify import CellHoriJustify
from ooo.dyn.table.cell_vert_justify import CellVertJustify


class BuildTable:
    HEADER_STYLE_NAME = "My HeaderStyle"
    DATA_STYLE_NAME = "My DataStyle"

    def __init__(self, im_fnm: PathOrStr, out_fnm: PathOrStr, **kwargs) -> None:
        _ = FileIO.is_exist_file(im_fnm, True)
        self._im_fnm = FileIO.get_absolute_path(im_fnm)
        if out_fnm:
            outf = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(outf)
            self._out_fnm = outf
        else:
            self._out_fnm = ""
        self._add_pic = bool(kwargs.get("add_pic", False))
        self._add_chart = bool(kwargs.get("add_chart", False))
        self._add_style = bool(kwargs.get("add_style", True))

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = Calc.create_doc(loader)

            GUI.set_visible(is_visible=True, odoc=doc)

            sheet = Calc.get_sheet(doc=doc, index=0)

            self._convert_addresses(sheet)

            # other possible build methods
            # self._buld_cells(sheet)
            # self._build_rows(sheet)
            # self._build_cols(sheet)

            self._build_array(sheet)

            if self._add_pic:
                self._add_picture(sheet=sheet, doc=doc)

            # add a chart
            if self._add_chart and Chart2:
                # assumes _build_array() has filled the spreadsheet with data
                rng_addr = Calc.get_address(sheet=sheet, range_name="B2:M4")
                chart_cell = "B6" if self._add_pic else "D6"
                Chart2.insert_chart(
                    sheet=sheet, cells_range=rng_addr, cell_name=chart_cell, width=21, height=11, diagram_name="Column"
                )

            if self._add_style:
                self._create_styles(doc)
                self._apply_styles(sheet)

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

    # region Private Methods

    def _buld_cells(self, sheet: XSpreadsheet) -> None:
        # column -- row
        header_vals = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
        for i, val in enumerate(header_vals):
            Calc.set_val(value=val, sheet=sheet, col=i + 1, row=0)

        # name
        vals = (31.45, 20.9, 117.5, 23.4, 114.5, 115.3, 171.3, 89.5, 41.2, 71.3, 25.4, 38.5)
        # start at B2
        for i, val in enumerate(vals):
            cell_name = TableHelper.make_cell_name(row=2, col=i + 2)
            Calc.set_val(value=val, sheet=sheet, cell_name=cell_name)

    def _build_rows(self, sheet: XSpreadsheet) -> None:
        vals = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
        Calc.set_row(sheet=sheet, values=vals, cell_name="B1")
        Calc.set_val(value="SUM", sheet=sheet, cell_name="N1")

        Calc.set_val(value="Smith", sheet=sheet, cell_name="A2")
        vals = (42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5)
        Calc.set_row(sheet=sheet, values=vals, cell_name="B2")
        Calc.set_val(value="=SUM(B2:M2)", sheet=sheet, cell_name="N2")

        Calc.set_val(value="Jones", sheet=sheet, col=0, row=2)
        vals = (21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5)
        Calc.set_row(sheet=sheet, values=vals, col_start=1, row_start=2)
        Calc.set_val(value="=SUM(B3:M3)", sheet=sheet, col=13, row=2)

        Calc.set_val(value="Brown", sheet=sheet, col=0, row=3)
        vals = (31.45, -20.9, -117.5, 23.4, -114.5, 115.3, -171.3, 89.5, 41.2, 71.3, 25.4, 38.5)
        Calc.set_row(sheet=sheet, values=vals, col_start=1, row_start=3)
        Calc.set_val(value="=SUM(A4:L4)", sheet=sheet, col=13, row=3)

    def _build_cols(self, sheet: XSpreadsheet) -> None:
        vals = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
        Calc.set_col(sheet=sheet, values=vals, cell_name="A2")
        Calc.set_val(value="SUM", sheet=sheet, cell_name="A14")

        Calc.set_val(value="Smith", sheet=sheet, cell_name="B1")
        vals = (42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5)
        Calc.set_col(sheet=sheet, values=vals, cell_name="B2")
        Calc.set_val(value="=SUM(B2:M2)", sheet=sheet, cell_name="B14")

        Calc.set_val(value="Jones", sheet=sheet, col=2, row=0)
        vals = (21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5)
        Calc.set_col(sheet=sheet, values=vals, col_start=2, row_start=1)
        Calc.set_val(value="=SUM(B3:M3)", sheet=sheet, col=2, row=13)

        Calc.set_val(value="Brown", sheet=sheet, col=3, row=0)
        vals = (31.45, -20.9, -117.5, 23.4, -114.5, 115.3, -171.3, 89.5, 41.2, 71.3, 25.4, 38.5)
        Calc.set_col(sheet=sheet, values=vals, col_start=3, row_start=1)
        Calc.set_val(value="=SUM(A4:L4)", sheet=sheet, col=3, row=13)

    def _build_array(self, sheet: XSpreadsheet) -> None:
        vals = (
            ("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"),
            ("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5),
            ("Jones", 21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5),
            ("Brown", 31.45, -20.9, -117.5, 23.4, -114.5, 115.3, -171.3, 89.5, 41.2, 71.3, 25.4, 38.5),
        )
        Calc.set_array(values=vals, sheet=sheet, name="A1:M4")  # or just A1

        Calc.set_val(sheet=sheet, cell_name="N1", value="SUM")
        Calc.set_val(sheet=sheet, cell_name="N2", value="=SUM(B2:M2)")
        Calc.set_val(sheet=sheet, cell_name="N3", value="=SUM(B3:M3)")
        Calc.set_val(sheet=sheet, cell_name="N4", value="=SUM(B4:M4)")

    def _convert_addresses(self, sheet: XSpreadsheet) -> None:
        # cell name <--> position
        pos = Calc.get_cell_position(cell_name="AA2")
        print(f"Positon of AA2: ({pos.X}, {pos.Y})")

        cell = Calc.get_cell(sheet=sheet, col=pos.X, row=pos.Y)
        Calc.print_cell_address(cell)

        print(f"AA2: {Calc.get_cell_str(col=pos.X, row=pos.Y)}")
        print()

        # cell range name <--> position
        rng = Calc.get_cell_range_positions("A1:D5")
        print(f"Range of A1:D5: ({rng[0].X}, {rng[0].Y}) -- ({rng[1].X}, {rng[1].Y})")

        cell_rng = Calc.get_cell_range(
            sheet=sheet, col_start=rng[0].X, row_start=rng[0].Y, col_end=rng[1].X, row_end=rng[1].Y
        )
        Calc.print_address(cell_rng)
        print(
            "A1:D5: " + Calc.get_range_str(col_start=rng[0].X, row_start=rng[0].Y, col_end=rng[1].X, row_end=rng[1].Y)
        )
        print()

    def _add_picture(self, sheet: XSpreadsheet, doc: XSpreadsheetDocument) -> None:
        # add a picture to the draw page for this sheet
        dp_sup = Lo.qi(XDrawPageSupplier, sheet, True)
        page = dp_sup.getDrawPage()
        x = 230 if self._add_chart else 125
        Draw.draw_image(slide=page, fnm=self._im_fnm, x=x, y=32)

        # look at all the draw pages
        supplier = Lo.qi(XDrawPagesSupplier, doc, True)
        pages = supplier.getDrawPages()
        print(f"1. No. of draw pages: {pages.getCount()}")

        comp_doc = Lo.qi(XComponent, doc, True)
        print(f"2. No. of draw pages: {Draw.get_slides_count(comp_doc)}")

    def _create_styles(self, doc: XSpreadsheetDocument) -> None:
        try:
            style1 = Calc.create_cell_style(doc=doc, style_name=BuildTable.HEADER_STYLE_NAME)

            Props.set(
                style1,
                IsCellBackgroundTransparent=False,
                CellBackColor=CommonColor.ROYAL_BLUE,
                CharColor=CommonColor.WHITE,
                HoriJustify=CellHoriJustify.CENTER,
                VertJustify=CellVertJustify.CENTER,
            )

            style2 = Calc.create_cell_style(doc=doc, style_name=BuildTable.DATA_STYLE_NAME)
            Props.set(
                style2,
                IsCellBackgroundTransparent=False,
                CellBackColor=CommonColor.LIGHT_BLUE,
                ParaRightMargin=500,  # move away from right edge
            )
        except Exception as e:
            print(e)

    def _apply_styles(sefl, sheet: XSpreadsheet) -> None:
        # apply cell styles
        # Calc.change_style(
        #     sheet=sheet, style_name=BuildTable.HEADER_STYLE_NAME, start_col=1, start_row=0, end_col=13, end_row=0
        # )
        # Calc.change_style(
        #     sheet=sheet, style_name=BuildTable.HEADER_STYLE_NAME, start_col=0, start_row=1, end_col=0, end_row=3
        # )
        # Calc.change_style(
        #     sheet=sheet, style_name=BuildTable.DATA_STYLE_NAME, start_col=1, start_row=1, end_col=13, end_row=3
        # )

        Calc.change_style(sheet=sheet, style_name=BuildTable.HEADER_STYLE_NAME, range_name="B1:N1")
        Calc.change_style(sheet=sheet, style_name=BuildTable.HEADER_STYLE_NAME, range_name="A2:A4")
        Calc.change_style(sheet=sheet, style_name=BuildTable.DATA_STYLE_NAME, range_name="B2:N4")

        Calc.add_border(
            sheet=sheet, range_name="A4:N4", color=CommonColor.DARK_BLUE, border_vals=Calc.BorderEnum.BOTTOM_BORDER
        )
        Calc.add_border(
            sheet=sheet,
            range_name="N1:N4",
            color=CommonColor.DARK_BLUE,
            border_vals=Calc.BorderEnum.LEFT_BORDER | Calc.BorderEnum.RIGHT_BORDER,
        )

    # endregion Private Methods

    # region properties
    @property
    def add_pic(self) -> bool:
        return self._add_pic

    @add_pic.setter
    def add_pic(self, value: bool):
        self._add_pic = value

    @property
    def add_chart(self) -> bool:
        return self._add_chart

    @add_chart.setter
    def add_chart(self, value: bool):
        self._add_chart = value

    @property
    def add_style(self) -> bool:
        return self._add_style
    
    @add_style.setter
    def add_style(self, value: bool):
        self._add_style = value
    # endregion properties
