#!/usr/bin/env python
from typing import cast, TYPE_CHECKING
import re
import contextlib
import uno
from ooo.dyn.ui.context_menu_interceptor_action import (
    ContextMenuInterceptorAction as ContextMenuAction,
)
from ooodev.adapter.ui.context_menu_interceptor import ContextMenuInterceptor
from ooodev.adapter.ui.context_menu_interceptor_event_data import (
    ContextMenuInterceptorEventData,
)
from ooodev.calc import CalcDoc, CalcCell, CalcSheetView

from ooodev.events.args.event_args_generic import EventArgsGeneric
from ooodev.gui.menu.context.action_trigger_item import ActionTriggerItem
from ooodev.gui.menu.context.action_trigger_sep import ActionTriggerSep
from .dispatch_provider_interceptor import DispatchProviderInterceptor


if TYPE_CHECKING:
    from com.sun.star.table import CellAddress


def on_menu_intercept(
    src: ContextMenuInterceptor,
    event: EventArgsGeneric[ContextMenuInterceptorEventData],
    view: CalcSheetView,
) -> None:
    """Event Listener for the context menu intercept."""
    with contextlib.suppress(Exception):
        # don't block other menus if there is an issue.

        container = event.event_data.event.action_trigger_container

        # The default action is ContextMenuAction.IGNORED.
        # In Linux as least (Ubuntu 20.04) the default will crash LibreOffice.
        # This is not the case in Windows, so this seems to be a bug.
        # Setting the action to ContextMenuAction.CONTINUE_MODIFIED will prevent the crash and
        # all seems to work well.
        event.event_data.action = ContextMenuAction.CONTINUE_MODIFIED

        # check the first and last items in the container
        if (
            container[0].CommandURL == ".uno:Cut"
            and container[-1].CommandURL == ".uno:FormatCellDialog"
        ):
            # get the current selection
            selection = event.event_data.event.selection.get_selection()

            if selection.getImplementationName() == "ScCellObj":
                # current selection is a cell.
                addr = cast("CellAddress", selection.getCellAddress())
                doc = CalcDoc.from_current_doc()
                sheet = doc.get_active_sheet()
                cell_obj = doc.range_converter.get_cell_obj_from_addr(addr)
                cell = sheet[cell_obj]
                if not is_http_url(cell.get_string()):
                    return

                if not has_url(cell):
                    # insert a new menu item.
                    # it's command is a custom .uno: command. Note that the command is not a built-in command.
                    # The command also contains args for the sheet and cell.
                    # A custom dispatch interceptor will be used to handle the command.
                    container.insert_by_index(4, ActionTriggerItem(f".uno:ooodev.calc.menu.convert.url?sheet={sheet.name}&cell={cell_obj}", "Convert to URL"))  # type: ignore
                    container.insert_by_index(4, ActionTriggerSep())  # type: ignore
                    event.event_data.action = ContextMenuAction.CONTINUE_MODIFIED


def is_http_url(s: str) -> bool:
    """Gets is a string matches a url pattern using regex"""
    url_pattern = re.compile(
        r"http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+"
    )
    return bool(re.match(url_pattern, s))


def has_url(cell: CalcCell) -> bool:
    """
    Gets if the cell contains a url.

    Args:
        cell (CalcCell): Calc Cell

    Returns:
        bool: True if the cell contains an actual linked url; Otherwise, False.
    """
    fields = cell.component.getTextFields()
    if len(fields) > 0:
        first = fields[0]
        ps = first.getPropertySetInfo()
        return ps.hasPropertyByName("URL")
    return False


def register_interceptor(doc: CalcDoc):
    """
    Registers the dispatch provider interceptor.

    This interceptor will be used to handle the custom .uno: command.

    Args:
        doc (CalcDoc): Calc Document
    """
    if DispatchProviderInterceptor.has_instance():
        return
    inst = DispatchProviderInterceptor()  # singleton
    frame = doc.get_frame()
    frame.registerDispatchProviderInterceptor(inst)
    view = doc.get_view()
    view.add_event_notify_context_menu_execute(on_menu_intercept)


def unregister_interceptor(doc: CalcDoc):
    """
    Un-registers the dispatch provider interceptor.

    Args:
        doc (CalcDoc): Calc Document
    """
    if not DispatchProviderInterceptor.has_instance():
        return
    inst = DispatchProviderInterceptor()  # singleton
    frame = doc.get_frame()
    frame.releaseDispatchProviderInterceptor(inst)
    view = doc.get_view()
    view.remove_event_notify_context_menu_execute(on_menu_intercept)
    inst.dispose()
