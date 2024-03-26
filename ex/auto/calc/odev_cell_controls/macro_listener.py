# region Imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno
from com.sun.star.awt import Key
from ooodev.calc import CalcDoc
from ooodev.form.controls.from_control_factory import FormControlFactory
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.utils.data_type.cell_obj import CellObj
from ooodev.dialog.msgbox import MessageBoxType


if TYPE_CHECKING:
    # design time only
    from com.sun.star.awt import ActionEvent
    from com.sun.star.drawing import ControlShape  # service
    from com.sun.star.table import CellAddress
    from com.sun.star.sheet import SheetCell  # service
    from com.sun.star.awt import KeyEvent  # struct
    from com.sun.star.sheet import ActivationEvent  # struct
    from ooodev.events.args.event_args import EventArgs
    from ooodev.calc.calc_cell import CalcCell
    from ooodev.calc.calc_sheet import CalcSheet
    from ooodev.form.controls.form_ctl_numeric_field import FormCtlNumericField
    from ooodev.form.controls.form_ctl_currency_field import FormCtlCurrencyField

# endregion Imports

# region Globals

_SELECTION_LISTENER = None
_CURRENT_DOC = None
_CURRENT_CELL = None
_PREV_CELL = None
_SHEET_INDEX = 0
_REMOVE_CTL_KEYS = (Key.ESCAPE, Key.RETURN, Key.TAB)
_AUTO_CONTROL = False

# endregion Globals

# region Doc


def get_current_doc() -> CalcDoc:
    global _CURRENT_DOC
    if _CURRENT_DOC is None:
        _CURRENT_DOC = CalcDoc.from_current_doc()
    return _CURRENT_DOC


# endregion Doc

# region Cost Column methods


def _process_cost_col(addr: CellAddress) -> None:
    global _PREV_CELL

    def _add_new_control(cell: CalcCell) -> None:
        ctl = cell.control.insert_control_currency_field(min_value=0, spin_button=True)
        ctl.model.Repeat = True
        ctl.value = cell.get_num()

        ctl.add_event_text_changed(on_cost_text_changed)
        ctl.add_event_key_released(on_cost_key_released)
        ctl.add_event_focus_lost(on_cost_focus_lost)

        ctl.view.setFocus()

    doc = get_current_doc()
    if addr.Column != 1:
        return
    if addr.Row < 1:
        return
    cell_obj = CellObj.from_cell(addr)
    sheet = doc.sheets.get_active_sheet()
    if _PREV_CELL is not None:
        try:
            previous_cell = sheet[_PREV_CELL]
            _remove_cost_ctl(sheet, previous_cell)
            previous_cell = None
        except Exception as e:
            print(f"Error Cleaning up previous cell: {e}")

    _add_new_control(sheet[cell_obj])


def _remove_cost_ctl(sheet: CalcSheet, cell: CalcCell) -> None:
    if cell.control.current_control is not None:
        ctl = cast("FormCtlCurrencyField", cell.control.current_control)
        ctl.remove_event_text_changed(on_cost_text_changed)
        ctl.remove_event_key_released(on_cost_key_released)
        # ctl.remove_event_focus_lost(on_cost_focus_lost)
        dp = sheet.draw_page
        shape = ctl.control_shape
        dp.remove(shape)
        ctl = None
        # print("Control removed")


# endregion Cost Column methods

# region Sold Column methods


def _process_sold_col(addr: CellAddress) -> None:
    global _PREV_CELL

    def _add_new_control(cell: CalcCell) -> None:
        ctl = cell.control.insert_control_numeric_field(min_value=0, spin_button=True)
        ctl.model.Repeat = True
        ctl.value = cell.get_num()
        ctl.add_event_text_changed(on_sold_text_changed)
        ctl.add_event_key_released(on_sold_key_released)
        ctl.add_event_focus_lost(on_sold_focus_lost)
        ctl.view.setFocus()

    doc = get_current_doc()
    if addr.Column != 2:
        return
    if addr.Row < 1:
        return
    cell_obj = CellObj.from_cell(addr)
    sheet = doc.sheets.get_active_sheet()
    if _PREV_CELL is not None:
        try:
            previous_cell = sheet[_PREV_CELL]
            _remove_sold_ctl(sheet, previous_cell)
            previous_cell = None
        except Exception as e:
            print(f"Error Cleaning up previous cell: {e}")

    _add_new_control(sheet[cell_obj])


def _remove_sold_ctl(sheet: CalcSheet, cell: CalcCell) -> None:
    if cell.control.current_control is not None:
        ctl = cast("FormCtlNumericField", cell.control.current_control)
        ctl.remove_event_text_changed(on_sold_text_changed)
        ctl.remove_event_key_released(on_sold_key_released)
        ctl.remove_event_focus_lost(on_sold_focus_lost)
        dp = sheet.draw_page
        shape = ctl.control_shape
        dp.remove(shape)
        ctl = None
        # print("Control removed")


# endregion Sold Column methods

# region Event Handlers

# region Start Stop Auto Control


def start_auto_control(*args) -> None:
    global _SELECTION_LISTENER, _AUTO_CONTROL
    doc = get_current_doc()
    if _AUTO_CONTROL:
        doc.msgbox(
            "Auto Control already started.", "Info", boxtype=MessageBoxType.WARNINGBOX
        )
        return
    _SELECTION_LISTENER = SelectionChangeEvents(doc=doc.component)
    _SELECTION_LISTENER.add_event_selection_changed(on_selection_changed)

    # remove and add the event listener that detects when the active sheet changes
    doc.current_controller.remove_event_active_spreadsheet_changed(
        on_active_sheet_changed
    )
    doc.current_controller.add_event_active_spreadsheet_changed(on_active_sheet_changed)
    _AUTO_CONTROL = True
    doc.msgbox("Auto Control is started", "Info", boxtype=MessageBoxType.INFOBOX)


def stop_auto_control(*args) -> None:
    global _SELECTION_LISTENER, _AUTO_CONTROL
    doc = get_current_doc()
    if not _AUTO_CONTROL:
        doc.msgbox(
            "Auto Control already stopped.", "Info", boxtype=MessageBoxType.WARNINGBOX
        )
        return
    if _SELECTION_LISTENER is not None:
        _SELECTION_LISTENER.remove_event_selection_changed(on_selection_changed)
        _SELECTION_LISTENER = None
    doc.current_controller.remove_event_active_spreadsheet_changed(
        on_active_sheet_changed
    )
    _AUTO_CONTROL = False
    doc.msgbox("Auto Control is stopped", "Info", boxtype=MessageBoxType.INFOBOX)


# endregion Start Stop Auto Control

# region Selection Changed Event


def on_selection_changed(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
    global _CURRENT_CELL, _PREV_CELL, _SHEET_INDEX
    doc = get_current_doc()
    sel = doc.get_current_selection()
    if sel is None:
        print("No selection")
        return
    keep_going = False
    with contextlib.suppress(Exception):
        # check for XCell
        if sel.ImplementationName == "ScCellObj":
            keep_going = True
    if not keep_going:
        return
    cell = cast("SheetCell", sel)
    addr = cell.getCellAddress()
    if addr.Sheet != _SHEET_INDEX:
        # only work with the first sheet
        # This not the most robust way to check for the first sheet
        # if the sheet moved or this will fail
        # If we went by sheet name, it may be a bit more robust but not foolproof
        # A better indicator should be used such as embedded a hidden value in the sheet or Hidden control in the sheet.
        return
    cell_obj = CellObj.from_cell(addr)
    _PREV_CELL = _CURRENT_CELL
    _CURRENT_CELL = cell_obj
    if _PREV_CELL == _CURRENT_CELL:
        return
    _process_sold_col(addr)
    _process_cost_col(addr)


# endregion Selection Changed Event


def on_text_changed(*args) -> None:
    pass


# region Button Event Handlers


def on_btn_action_preformed(event: ActionEvent) -> None:
    doc = get_current_doc()
    sheet = doc.sheets.get_active_sheet()
    factory = FormControlFactory(
        draw_page=sheet.draw_page.component, lo_inst=doc.lo_inst
    )
    shape = cast("ControlShape", factory.find_shape_for_control(event.Source.Model))

    # the anchor for the control shape is the cell
    x_cell = shape.Anchor
    # get a CellOjb instance that is easier to work with
    cell_obj = doc.range_converter.get_cell_obj(x_cell)
    cell_end = cell_obj.left
    row_index = cell_end.row_obj.index
    # create a RangeObj for the row data
    row_rng = doc.range_converter.rng_from_position(
        0, row_index, cell_end.col_obj.index, row_index, sheet.sheet_index
    )
    data = sheet.get_array(range_obj=row_rng)
    doc.msgbox(f"Data: {data[0]}", "Info")


# endregion Button Event Handlers

# region Cost Column Event Handlers


def on_cost_text_changed(
    src: Any, event: EventArgs, control_src: FormCtlCurrencyField, *args, **kwargs
) -> None:
    doc = get_current_doc()
    sheet = doc.sheets.get_active_sheet()
    x_cell = control_src.control_shape.Anchor
    cell_obj = doc.range_converter.get_cell_obj(x_cell)
    sheet[cell_obj].value = control_src.value


def on_cost_key_released(
    src: Any, event: EventArgs, control_src: FormCtlCurrencyField, *args, **kwargs
) -> None:
    global _REMOVE_CTL_KEYS
    kc = cast("KeyEvent", event.event_data)
    if kc.KeyCode in _REMOVE_CTL_KEYS:
        doc = get_current_doc()
        sheet = doc.sheets.get_active_sheet()
        x_cell = control_src.control_shape.Anchor
        _remove_cost_ctl(sheet, sheet.get_cell(x_cell))


def on_cost_focus_lost(
    src: Any, event: EventArgs, control_src: FormCtlCurrencyField, *args, **kwargs
) -> None:
    doc = get_current_doc()
    sheet = doc.sheets.get_active_sheet()
    x_cell = control_src.control_shape.Anchor
    _remove_cost_ctl(sheet, sheet.get_cell(x_cell))


# endregion Cost Column Event Handlers

# region Sold Column Event Handlers


def on_sold_text_changed(
    src: Any, event: EventArgs, control_src: FormCtlNumericField, *args, **kwargs
) -> None:
    doc = get_current_doc()
    sheet = doc.sheets.get_active_sheet()
    x_cell = control_src.control_shape.Anchor
    cell_obj = doc.range_converter.get_cell_obj(x_cell)
    sheet[cell_obj].value = control_src.value


def on_sold_key_released(
    src: Any, event: EventArgs, control_src: FormCtlNumericField, *args, **kwargs
) -> None:
    global _REMOVE_CTL_KEYS
    kc = cast("KeyEvent", event.event_data)
    # print("Key code", kc.KeyCode)
    # print("Key char", kc.KeyChar)
    if kc.KeyCode in _REMOVE_CTL_KEYS:
        doc = get_current_doc()
        sheet = doc.sheets.get_active_sheet()
        x_cell = control_src.control_shape.Anchor
        _remove_sold_ctl(sheet, sheet.get_cell(x_cell))


def on_sold_focus_lost(
    src: Any, event: EventArgs, control_src: FormCtlNumericField, *args, **kwargs
) -> None:
    doc = get_current_doc()
    sheet = doc.sheets.get_active_sheet()
    x_cell = control_src.control_shape.Anchor
    _remove_sold_ctl(sheet, sheet.get_cell(x_cell))


# endregion Sold Column Event Handlers


# region Sheet Activation Event
def on_active_sheet_changed(
    source: Any, event_args: EventArgs, *args, **kwargs
) -> None:
    # see the ex/auto/calc/odev_sheet_activation_event/runner.py for an example
    # There is a bug with Calc that causes controls to not work properly when the sheet is changed
    # See: https://bugs.documentfoundation.org/show_bug.cgi?id=159134
    # By monitoring the active sheet change event, we can remove the control from the cell if it exists
    # Although this even happens after the sheet has be changed, the issue is not noticeable until the user changes back to the sheet with the control.
    # At the point when the sheet with the control is activated, the control will not work and possibly cause a crash.
    # This is in only an issue when the controls are added to the sheet programmatically not using the ctl.assign_script method.
    # The buttons added to the sheet using ctl.assign_script method will work as expected.
    global _CURRENT_CELL, _SHEET_INDEX, _PREV_CELL
    print("Active Sheet Changed")
    doc = get_current_doc()
    try:
        if _CURRENT_CELL is not None:
            event = cast("ActivationEvent", event_args.event_data)
            sheet = doc.sheets.get_sheet(event.ActiveSheet)
            # print("    Active Sheet:", sheet.name)
            if sheet.sheet_index == _SHEET_INDEX:
                cell = sheet[_CURRENT_CELL]
                _remove_sold_ctl(sheet, cell)
                _remove_cost_ctl(sheet, cell)
                # _CURRENT_CELL = None
                # _PREV_CELL = None
    except Exception as e:
        print(f"  {e}")


# endregion Sheet Activation Event

# endregion Event Handlers

# region Other Macros


def info(*args) -> None:
    # just a sample function for testing
    doc = get_current_doc()
    doc.msgbox("Hello from Python", "Info")


# endregion Other Macros

g_exportedScripts = (
    info,
    on_btn_action_preformed,
    start_auto_control,
    stop_auto_control,
)
