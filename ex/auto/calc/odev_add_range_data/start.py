#!/usr/bin/env python
from __future__ import annotations
from pathlib import Path
import uno
from ooo.dyn.frame.infobar_type import InfobarTypeEnum

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.loader import Lo
from ooodev.calc import CalcDoc, CalcSheet, ZoomKind, CalcSheetView
from ooodev.utils.color import CommonColor
from ooodev.adapter.util.the_path_settings_comp import ThePathSettingsComp
from ooodev.utils.props import Props


def do_cell_range(sheet: CalcSheet) -> None:
    vals = (
        ("Name", "Fruit", "Quantity"),
        ("Alice", "Apples", 3),
        ("Alice", "Oranges", 7),
        ("Bob", "Apples", 3),
        ("Alice", "Apples", 9),
        ("Bob", "Apples", 5),
        ("Bob", "Oranges", 6),
        ("Alice", "Oranges", 3),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 1),
        ("Bob", "Oranges", 2),
        ("Bob", "Oranges", 7),
        ("Bob", "Apples", 1),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 8),
        ("Alice", "Apples", 7),
        ("Bob", "Apples", 1),
        ("Bob", "Oranges", 9),
        ("Bob", "Oranges", 3),
        ("Alice", "Oranges", 4),
        ("Alice", "Apples", 9),
    )
    with sheet.calc_doc:
        # use doc context manager to lock controllers for faster updates.
        sheet.set_array(values=vals, name="A1:C21")  # or just "A1"
        cell = sheet.get_cell(cell_name="A22")
        cell.set_val("Total")

        cell = sheet.get_cell(cell_name="C22")
        cell.set_val("=SUM(C2:C21)")

        # set Border around data and summary.
        rng = sheet.get_range(range_name="A1:C22")
        rng.style_borders_sides(color=CommonColor.LIGHT_BLUE, width=2.85)


def get_docs_dir() -> Path:
    paths = ThePathSettingsComp.from_lo()
    return Path(uno.fileUrlToSystemPath(paths.work[0]))


def add_info_bar_doc_saved(doc: CalcDoc, path: str):
    # Create a new info bar
    view = doc.get_view()
    buttons = (Props.make_sting_pair("Close doc", ".uno:CloseDoc"),)
    view.append_infobar(
        id="DocSaved",
        primary_message="Document saved as:",
        secondary_message=f"Path: {path}",
        infobar_type=InfobarTypeEnum.INFO,
        action_buttons=buttons,
        show_close_button=True,
    )


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = CalcDoc.create_doc(visible=True)
        sheet = doc.sheets[0]
        # doc.set_visible()
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.PAGE_WIDTH)

        do_cell_range(sheet=sheet)
        doc.freeze_rows(num_rows=1)

        Lo.delay(1_500)
        msg_result = doc.msgbox(
            "Do you wish to save the document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            doc_path = get_docs_dir() / "odev_add_range_data.ods"
            doc.save_doc(doc_path)
            add_info_bar_doc_saved(doc=doc, path=str(doc_path))
        else:

            msg_result = doc.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.close_doc()
                Lo.close_office()
                return

        print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
