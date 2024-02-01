from __future__ import annotations
from pathlib import Path
import tempfile

import uno

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import CalcDoc
from ooodev.draw import ImpressDoc
from ooodev.exceptions import ex as mEx
from ooodev.office.chart2 import Chart2, Angle
from ooodev.office.draw import DrawingNameSpaceKind
from ooodev.units import UnitMM
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.dispatch.global_edit_dispatch import GlobalEditDispatch
from ooodev.utils.file_io import FileIO
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class SlideChart:
    def __init__(self, data_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(data_fnm, True)
        self._data_fnm = FileIO.get_absolute_path(data_fnm)
        self._out_dir = Path(tempfile.mkdtemp())

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            calc_doc = None
            try:
                calc_doc = self._make_col_chart()
            except Exception:
                calc_doc = None
            has_chart = calc_doc is not None

            doc = ImpressDoc.create_doc(loader=loader, visible=True)

            # access first page.
            slide = doc.slides[0]
            body = slide.bullets_slide(title="Sneakers Are Selling!")
            body.add_bullet(level=0, text="Sneaker profits have increased")

            if has_chart:
                doc.activate()
                Lo.delay(1_000)
                Lo.dispatch_cmd(GlobalEditDispatch.PASTE)

            try:
                ole_shape = slide.find_shape_by_type(
                    shape_type=DrawingNameSpaceKind.OLE2_SHAPE
                )
                slide_size = slide.get_size_mm()
                shape_size = ole_shape.get_size_mm()
                shape_pos = ole_shape.get_position_mm()

                y = slide_size.Height - shape_size.Height - 20
                # move pic down
                ole_shape.set_position(x=UnitMM(shape_pos.X), y=UnitMM(y))
            except mEx.ShapeMissingError:
                Lo.print("Did not find shape, unable to set size and position")

            Lo.delay(2000)
            msg_result = doc.msgbox(
                "Do you wish to close documents?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                if calc_doc is not None:
                    calc_doc.close()
                doc.close()
                Lo.close_office()
            else:
                print("Keeping document open")
        except Exception:
            Lo.close_office()
            raise

    def _make_col_chart(self) -> CalcDoc:
        lo_inst = Lo.create_lo_instance()
        ss_doc = CalcDoc.open_doc(fnm=self._data_fnm, lo_inst=lo_inst)
        ss_doc.set_visible(True)

        try:
            sheet = ss_doc.sheets[0]

            range_addr = sheet.get_address(range_name="A2:B8")

            # Switch context so Chart2 is using the same context as the CalcDoc
            with LoContext(lo_inst):
                chart_doc = Chart2.insert_chart(
                    sheet=sheet.component,
                    cells_range=range_addr,
                    cell_name="C3",
                    width=10,
                    height=8,
                    diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
                )

                Chart2.set_title(chart_doc=chart_doc, title=sheet["A1"].get_string())
                Chart2.set_x_axis_title(
                    chart_doc=chart_doc, title=sheet["A2"].get_string()
                )
                Chart2.set_y_axis_title(
                    chart_doc=chart_doc, title=sheet["B2"].get_string()
                )
                Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
                ss_doc.activate()
                Lo.delay(1_000)

                Chart2.copy_chart(ssdoc=ss_doc.component, sheet=sheet.component)

            try:
                # Switch context so ImagesLo is using the same context as the CalcDoc
                with LoContext(lo_inst):
                    ImagesLo.save_graphic(
                        pic=Chart2.get_chart_image(sheet.component),
                        fnm=Path(self._out_dir, "chartImage.png"),
                    )
            except mEx.ImageError:
                pass

        except Exception as e:
            lo_inst.print("Error making col chart")
            lo_inst.print(f"  {e}")
            raise
        return ss_doc
