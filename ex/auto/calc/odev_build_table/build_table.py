from __future__ import annotations

import uno
from com.sun.star.drawing import XDrawPagesSupplier
from com.sun.star.lang import XComponent

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.format import Styler
from ooodev.format.calc.direct.cell import borders as direct_borders
from ooodev.format.calc.modify.cell import borders as modify_borders
from ooodev.format.calc.modify.cell.alignment import (
    HoriAlignKind,
    VertAlignKind,
    TextAlign,
)
from ooodev.format.calc.modify.cell.background import Color as BgColor
from ooodev.format.calc.modify.cell.font import FontEffects
from ooodev.calc import Calc, CalcDoc, CalcSheet, ZoomKind
from ooodev.office.draw import Draw
from ooodev.units import UnitMM
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.type_var import PathOrStr

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    # bug in LibreOffice 7.4, Fixed in 7.5
    Chart2 = None


class BuildTable:
    HEADER_STYLE_NAME = "My HeaderStyle"
    DATA_STYLE_NAME = "My DataStyle"

    def __init__(self, im_fnm: PathOrStr, out_fnm: PathOrStr, **kwargs) -> None:
        _ = FileIO.is_exist_file(im_fnm, True)
        self._im_fnm = FileIO.get_absolute_path(im_fnm)
        if out_fnm:
            out_f = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(out_f)
            self._out_fnm = out_f
        else:
            self._out_fnm = ""
        self._add_pic = bool(kwargs.get("add_pic", False))
        self._add_chart = bool(kwargs.get("add_chart", False))
        self._add_style = bool(kwargs.get("add_style", True))

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = CalcDoc.create_doc(loader=loader, visible=True)

            Lo.delay(300)
            doc.zoom(ZoomKind.ZOOM_100_PERCENT)

            sheet = doc.sheets[0]

            self._convert_addresses(sheet)

            # other possible build methods
            # self._build_cells(sheet)
            # self._build_rows(sheet)
            # self._build_cols(sheet)

            self._build_array(sheet)

            if self._add_pic:
                self._add_picture(sheet)

            # add a chart
            if self._add_chart and Chart2:
                # assumes _build_array() has filled the spreadsheet with data
                rng_addr = sheet.get_address(range_name="B2:M4")
                chart_cell = "B6" if self._add_pic else "D6"
                Chart2.insert_chart(
                    sheet=sheet.component,
                    cells_range=rng_addr,
                    cell_name=chart_cell,
                    width=21,
                    height=11,
                    diagram_name="Column",
                )

            if self._add_style:
                self._create_styles(doc)
                self._apply_styles(sheet)

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

    # region Private Methods

    def _build_cells(self, sheet: CalcSheet) -> None:
        # column -- row
        header_vals = (
            "JAN",
            "FEB",
            "MAR",
            "APR",
            "MAY",
            "JUN",
            "JUL",
            "AUG",
            "SEP",
            "OCT",
            "NOV",
            "DEC",
        )

        for i, val in enumerate(header_vals):
            sheet.set_val(value=val, col=i + 1, row=0)

        # name
        vals = (
            31.45,
            20.9,
            117.5,
            23.4,
            114.5,
            115.3,
            171.3,
            89.5,
            41.2,
            71.3,
            25.4,
            38.5,
        )
        # start at B2
        for i, val in enumerate(vals):
            cell_name = TableHelper.make_cell_name(row=2, col=i + 2)
            sheet.set_val(value=val, cell_name=cell_name)

    def _build_rows(self, sheet: CalcSheet) -> None:
        vals = (
            "JAN",
            "FEB",
            "MAR",
            "APR",
            "MAY",
            "JUN",
            "JUL",
            "AUG",
            "SEP",
            "OCT",
            "NOV",
            "DEC",
        )
        sheet.set_row(values=vals, cell_name="B1")
        sheet.set_val(value="SUM", cell_name="N1")

        sheet.set_val(value="Smith", cell_name="A2")
        vals = (42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5)
        sheet.set_row(values=vals, cell_name="B2")
        sheet.set_val(value="=SUM(B2:M2)", cell_name="N2")

        sheet.set_val(value="Jones", col=0, row=2)
        vals = (21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5)
        sheet.set_row(values=vals, col_start=1, row_start=2)
        sheet.set_val(value="=SUM(B3:M3)", col=13, row=2)

        sheet.set_val(value="Brown", col=0, row=3)
        vals = (
            31.45,
            -20.9,
            -117.5,
            23.4,
            -114.5,
            115.3,
            -171.3,
            89.5,
            41.2,
            71.3,
            25.4,
            38.5,
        )
        sheet.set_row(values=vals, col_start=1, row_start=3)
        sheet.set_val(value="=SUM(A4:L4)", col=13, row=3)

    def _build_cols(self, sheet: CalcSheet) -> None:
        vals = (
            "JAN",
            "FEB",
            "MAR",
            "APR",
            "MAY",
            "JUN",
            "JUL",
            "AUG",
            "SEP",
            "OCT",
            "NOV",
            "DEC",
        )
        sheet.set_col(values=vals, cell_name="A2")
        sheet.set_val(value="SUM", cell_name="A14")

        sheet.set_val(value="Smith", cell_name="B1")
        vals = (42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5)
        sheet.set_col(values=vals, cell_name="B2")
        sheet.set_val(value="=SUM(B2:M2)", cell_name="B14")

        sheet.set_val(value="Jones", col=2, row=0)
        vals = (21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5)
        sheet.set_col(values=vals, col_start=2, row_start=1)
        sheet.set_val(value="=SUM(B3:M3)", col=2, row=13)

        sheet.set_val(value="Brown", col=3, row=0)
        vals = (
            31.45,
            -20.9,
            -117.5,
            23.4,
            -114.5,
            115.3,
            -171.3,
            89.5,
            41.2,
            71.3,
            25.4,
            38.5,
        )
        sheet.set_col(values=vals, col_start=3, row_start=1)
        sheet.set_val(value="=SUM(A4:L4)", col=3, row=13)

    def _build_array(self, sheet: CalcSheet) -> None:
        vals = (
            (
                "",
                "JAN",
                "FEB",
                "MAR",
                "APR",
                "MAY",
                "JUN",
                "JUL",
                "AUG",
                "SEP",
                "OCT",
                "NOV",
                "DEC",
            ),
            (
                "Smith",
                42,
                58.9,
                -66.5,
                43.4,
                44.5,
                45.3,
                -67.3,
                30.5,
                23.2,
                -97.3,
                22.4,
                23.5,
            ),
            (
                "Jones",
                21,
                40.9,
                -57.5,
                -23.4,
                34.5,
                59.3,
                27.3,
                -38.5,
                43.2,
                57.3,
                25.4,
                28.5,
            ),
            (
                "Brown",
                31.45,
                -20.9,
                -117.5,
                23.4,
                -114.5,
                115.3,
                -171.3,
                89.5,
                41.2,
                71.3,
                25.4,
                38.5,
            ),
        )
        sheet.set_array(values=vals, name="A1:M4")  # or just A1

        sheet.set_val(cell_name="N1", value="SUM")
        sheet.set_val(cell_name="N2", value="=SUM(B2:M2)")
        sheet.set_val(cell_name="N3", value="=SUM(B3:M3)")
        sheet.set_val(cell_name="N4", value="=SUM(B4:M4)")

    def _convert_addresses(self, sheet: CalcSheet) -> None:
        # cell name <--> position
        pos = sheet["A22"].get_cell_position()
        print(f"Position of AA2: ({pos.X}, {pos.Y})")

        cell = sheet[(pos.X, pos.Y)]
        Calc.print_cell_address(cell.component)

        print(f"AA2: {cell.get_cell_str()}")
        print()

        # cell range name <--> position
        rng = Calc.get_cell_range_positions("A1:D5")
        print(f"Range of A1:D5: ({rng[0].X}, {rng[0].Y}) -- ({rng[1].X}, {rng[1].Y})")

        cell_rng = Calc.get_cell_range(
            sheet=sheet.component,
            col_start=rng[0].X,
            row_start=rng[0].Y,
            col_end=rng[1].X,
            row_end=rng[1].Y,
        )
        Calc.print_address(cell_rng)
        print(
            "A1:D5: "
            + Calc.get_range_str(
                col_start=rng[0].X,
                row_start=rng[0].Y,
                col_end=rng[1].X,
                row_end=rng[1].Y,
            )
        )
        print()

    def _add_picture(self, sheet: CalcSheet) -> None:
        # add a picture to the draw page for this sheet
        x = 230 if self._add_chart else 125
        sheet.draw_page.draw_image(fnm=self._im_fnm, x=x, y=32)

        # look at all the draw pages
        # 3 ways to get the draw pages
        supplier = sheet.calc_doc.qi(XDrawPagesSupplier, True)
        pages = supplier.getDrawPages()
        print(f"1. No. of draw pages: {pages.getCount()}")

        comp_doc = sheet.calc_doc.qi(XComponent, True)
        print(f"2. No. of draw pages: {Draw.get_slides_count(comp_doc)}")

        print(f"3. No. of draw pages: {len(sheet.calc_doc.draw_pages)}")

    def _create_styles(self, doc: CalcDoc) -> None:
        try:
            # create a style using Calc
            header_style = doc.create_cell_style(
                style_name=BuildTable.HEADER_STYLE_NAME
            )

            # create formats to apply to header_style
            header_bg_color_style = BgColor(
                color=CommonColor.ROYAL_BLUE, style_name=BuildTable.HEADER_STYLE_NAME
            )
            effects_style = FontEffects(
                color=CommonColor.WHITE, style_name=BuildTable.HEADER_STYLE_NAME
            )
            txt_align_style = TextAlign(
                hori_align=HoriAlignKind.CENTER,
                vert_align=VertAlignKind.MIDDLE,
                style_name=BuildTable.HEADER_STYLE_NAME,
            )
            # Apply formatting to header_style
            Styler.apply(
                header_style, header_bg_color_style, effects_style, txt_align_style
            )

            # create style
            data_style = doc.create_cell_style(style_name=BuildTable.DATA_STYLE_NAME)

            # create formats to apply to data_style
            footer_bg_color_style = BgColor(
                color=CommonColor.LIGHT_BLUE, style_name=BuildTable.DATA_STYLE_NAME
            )
            bdr_style = modify_borders.Borders(
                padding=modify_borders.Padding(left=UnitMM(5))
            )

            # Apply formatting to data_style
            Styler.apply(data_style, footer_bg_color_style, bdr_style, txt_align_style)

        except Exception as e:
            print(e)

    def _apply_styles(self, sheet: CalcSheet) -> None:
        sheet.change_style(style_name=BuildTable.HEADER_STYLE_NAME, range_name="B1:N1")

        sheet.change_style(style_name=BuildTable.HEADER_STYLE_NAME, range_name="A2:A4")
        rng = sheet.get_range(range_name="B2:N4")
        rng.change_style(style_name=BuildTable.DATA_STYLE_NAME)

        # create a border side, default width units are points
        side = direct_borders.Side(width=2.85, color=CommonColor.DARK_BLUE)
        # create a border setting bottom side
        bdr = direct_borders.Borders(bottom=side)
        # Apply border to range

        sheet.set_style_range(range_name="A4:N4", styles=[bdr])

        # create a border with left and right
        bdr = direct_borders.Borders(left=side, right=side)
        # Apply border to range
        rng = sheet.get_range(range_name="N1:N4")
        rng.set_style(styles=[bdr])

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
