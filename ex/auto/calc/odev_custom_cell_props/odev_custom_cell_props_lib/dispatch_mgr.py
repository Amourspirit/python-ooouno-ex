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

from ooodev.io.log import logging as logger
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
                if not cell.has_custom_property("OriginalValue"):
                    logger.debug(
                        f"Cell {cell_obj} does not have OriginalValue custom property."
                    )
                    return
                if is_cell_orig_value(cell):
                    logger.debug(f"Cell {cell_obj} is already at its original value.")
                    return
                # insert a new menu item.
                # it's command is a custom .uno: command. Note that the command is not a built-in command.
                # The command also contains args for the sheet and cell.
                # A custom dispatch interceptor will be used to handle the command.

                logger.debug(
                    f"Cell {cell_obj} value of '{cell.value}' is not at its original value of '{cell.get_custom_property('OriginalValue')}'"
                )
                container.insert_by_index(4, ActionTriggerItem(f".uno:ooodev.calc.menu.update.orig?sheet={sheet.name}&cell={cell_obj}", "Update Original"))  # type: ignore
                container.insert_by_index(4, ActionTriggerItem(f".uno:ooodev.calc.menu.reset.orig?sheet={sheet.name}&cell={cell_obj}", "Rest to Original"))  # type: ignore
                container.insert_by_index(4, ActionTriggerSep())  # type: ignore
                event.event_data.action = ContextMenuAction.CONTINUE_MODIFIED


def float_almost_equal(a, b, epsilon=1e-9):
    return abs(a - b) < epsilon


def is_cell_orig_value(cell: CalcCell) -> bool:
    """
    Compares the value of a cell with its custom property OriginalValue.

    Args:
        cell (CalcCell): Cell to compare.

    Returns:
        bool: True if the value is different from the custom property OriginalValue.
    """
    if cell.value is None:
        return False
    current_val = cell.value
    original_val = cell.get_custom_property("OriginalValue")
    if type(current_val) is not type(original_val):
        return False
    if isinstance(current_val, float) and isinstance(original_val, (float, int)):
        return float_almost_equal(current_val, original_val)
    return current_val == original_val


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
