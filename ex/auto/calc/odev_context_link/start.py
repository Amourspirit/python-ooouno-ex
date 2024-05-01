#!/usr/bin/env python
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import re
import uno
from ooo.dyn.ui.context_menu_interceptor_action import (
    ContextMenuInterceptorAction as ContextMenuAction,
)
from ooodev.adapter.ui.context_menu_interceptor import ContextMenuInterceptor
from ooodev.adapter.ui.context_menu_interceptor_event_data import (
    ContextMenuInterceptorEventData,
)
from ooodev.calc import CalcDoc, CalcSheet, CalcCell, CalcSheetView, ZoomKind
from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.events.args.event_args_generic import EventArgsGeneric
from ooodev.gui.menu.context.action_trigger_item import ActionTriggerItem
from ooodev.loader import Lo
from dispatch_provider_interceptor import DispatchProviderInterceptor


if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from ooodev.events.args.event_args import EventArgs


INTERCEPTOR = None


def find_http_urls(text):
    url_pattern = re.compile(
        r"http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+"
    )
    urls = re.findall(url_pattern, text)
    return urls


def is_http_url(s: str) -> bool:
    url_pattern = re.compile(
        r"http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+"
    )
    return bool(re.match(url_pattern, s))


def on_menu_intercept(
    src: ContextMenuInterceptor,
    event: EventArgsGeneric[ContextMenuInterceptorEventData],
    view: CalcSheetView,
) -> None:
    try:
        container = event.event_data.event.action_trigger_container
        # don't block other menus if there is an issue.
        # check the first and last items in the container
        # print(container.get_by_index(0).CommandURL)
        event.event_data.action = ContextMenuAction.CONTINUE_MODIFIED
        if (
            container[0].CommandURL == ".uno:Cut"
            and container[-1].CommandURL == ".uno:FormatCellDialog"
        ):

            selection = event.event_data.event.selection.get_selection()
            # print(selection)

            if selection.getImplementationName() == "ScCellObj":
                # print(dir(selection))
                # addr = cast(CellAddress, selection.CellAddress)
                addr = cast("CellAddress", selection.getCellAddress())
                doc = CalcDoc.from_current_doc()
                sheet = doc.get_active_sheet()
                cell_obj = doc.range_converter.get_cell_obj_from_addr(addr)
                cell = sheet[cell_obj]
                if not is_http_url(cell.get_string()):
                    return

                if not has_url(cell):

                    # url = MacroScript.get_url_script(
                    #     name="convert_cell_url",
                    #     library="convert_url",
                    #     language="Python",
                    #     location="user",
                    # )
                    container.insert_by_index(1, ActionTriggerItem(f".uno:ooodev.calc.menu.convert.url?sheet={sheet.name}&cell={cell_obj}", "Convert to URL"))  # type: ignore
                    event.event_data.action = ContextMenuAction.CONTINUE_MODIFIED

    except Exception as e:
        print(e)
        raise


def set_cell_data(sheet: CalcSheet) -> None:
    vals = (
        ("Hyperlinks",),
        (
            "https://ask.libreoffice.org/t/how-to-convert-links-into-hyperlinks-in-bulk-in-calc/102448",
        ),
        (
            "https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets",
        ),
        (
            "https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_cell_texts",
        ),
        (
            "https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/data_type/range_obj.html",
        ),
        ("https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html",),
        (
            "https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/index.html",
        ),
    )
    sheet.set_array(values=vals, name="A1")
    # bold the header
    _ = sheet["A1"].style_font_general(b=True)
    col = sheet.get_col_range(0)
    col.optimal_width = True


def create_hyperlinks() -> None:
    # get access to current Calc Document
    doc = CalcDoc.from_current_doc()

    # get access to first spreadsheet
    sheet = doc.sheets[0]

    # insert the array of data
    set_cell_data(sheet=sheet)


def convert_to_hyperlink(cell: CalcCell):
    cell_data = cell.get_string()
    if cell_data.startswith("http"):
        cell.value = ""
        cursor = cell.create_text_cursor()
        cursor.add_hyperlink(
            label=cell_data,
            url_str=cell_data,
        )


def has_url(cell: CalcCell) -> bool:
    fields = cell.component.getTextFields()
    if len(fields) > 0:
        first = fields[0]
        ps = first.getPropertySetInfo()
        return ps.hasPropertyByName("URL")
    return False


def register_interceptor(doc: CalcDoc):
    global INTERCEPTOR
    INTERCEPTOR = DispatchProviderInterceptor()
    frame = doc.get_frame()
    frame.registerDispatchProviderInterceptor(INTERCEPTOR)


def main() -> int:
    global INTERCEPTOR
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = CalcDoc.create_doc(visible=True)
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)
        register_interceptor(doc)

        create_hyperlinks()
        view = doc.get_view()
        view.add_event_notify_context_menu_execute(on_menu_intercept)
        # Lo.delay(1_500)
        doc.close_doc()
        Lo.close_office()
        return
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
    INTERCEPTOR = None
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
