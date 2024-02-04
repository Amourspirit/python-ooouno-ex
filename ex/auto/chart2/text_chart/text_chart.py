from __future__ import annotations

import uno
from com.sun.star.frame import XComponentLoader

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.loader import Lo
from ooodev.calc import CalcDoc
from ooodev.office.chart2 import Chart2, Angle
from ooodev.write import WriteDoc
from ooodev.utils.dispatch.global_edit_dispatch import GlobalEditDispatch
from ooodev.utils.file_io import FileIO
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.style.paragraph_adjust import ParagraphAdjust


class TextChart:
    def __init__(self, data_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(data_fnm, True)
        self._data_fnm = FileIO.get_absolute_path(data_fnm)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            has_chart = self._make_col_chart(loader)

            doc = WriteDoc.create_doc(loader=loader, visible=True)

            cursor = doc.get_cursor()
            # make sure at end of doc before appending
            cursor.goto_end()

            cursor.append_para("Hello LibreOffice.\n")

            if has_chart:
                Lo.delay(1_000)
                Lo.dispatch_cmd(GlobalEditDispatch.PASTE)

            cursor.append_para("Figure 1. Sneakers Column Chart.\n")
            cursor.style_prev_paragraph(
                prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust"
            )

            cursor.append_para("Some more text...\n")

            Lo.delay(2000)
            msg_result = doc.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.close()
                Lo.close_office()
            else:
                print("Keeping document open")
        except Exception:
            Lo.close_office()
            raise

    def _make_col_chart(self, loader: XComponentLoader) -> bool:
        # create a new context to use with second document
        lo_inst = Lo.create_lo_instance()
        doc = CalcDoc.open_doc(
            fnm=self._data_fnm, loader=loader, lo_inst=lo_inst, visible=True
        )
        try:
            sheet = doc.sheets[0]

            range_addr = sheet.get_address(range_name="A2:B8")
            chart_doc = Chart2.insert_chart(
                sheet=sheet.component,
                cells_range=range_addr,
                cell_name="C3",
                width=15,
                height=11,
                diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
            )
            Chart2.set_title(chart_doc=chart_doc, title=sheet["A1"].get_string())
            Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["A2"].get_string())
            Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["B2"].get_string())
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
            Lo.delay(1_000)
            Chart2.copy_chart(ssdoc=doc.component, sheet=sheet.component)
            return True
        except Exception as e:
            Lo.print("Error making col chart")
            Lo.print(f"  {e}")
        finally:
            doc.close()
        return False
