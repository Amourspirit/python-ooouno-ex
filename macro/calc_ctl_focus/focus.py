from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import os
from datetime import datetime
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooodev.calc import CalcDoc
from ooodev.events.args.event_args import EventArgs
from ooodev.form.controls.form_ctl_text_field import FormCtlTextField
from ooodev.utils.color import StandardColor

if TYPE_CHECKING:
    from com.sun.star.awt import FocusEvent
    from com.sun.star.awt import MouseEvent

_TEXT_OUTPUT = cast(FormCtlTextField, None)


def _write_line(msg: str):
    global _TEXT_OUTPUT
    if _TEXT_OUTPUT is not None:
        now = datetime.now()
        timestamp = now.strftime("%H:%M:%S")
        _TEXT_OUTPUT.write_line(f"{timestamp}: {msg}")


def on_lost_focus(
    src: Any, event: EventArgs, control_src: FormCtlTextField, *args, **kwargs
):
    _write_line("Lost Focus")
    fe = cast("FocusEvent", event.event_data)
    if fe.Temporary is False:
        if fe.NextFocus is not None:
            if (
                fe.NextFocus.PosSize.X != fe.Source.PosSize.X
                or fe.NextFocus.PosSize.Y != fe.Source.PosSize.Y
            ):
                _write_line("  The control can be deleted")
                return
    _write_line("  The control cannot be deleted:")


def on_mouse_event(src: Any, event: EventArgs, *args, **kwargs):
    _write_line("on_mouse_event")
    me = cast("MouseEvent", event.event_data)
    _write_line(f"  PopupTrigger: {me.PopupTrigger}")


def on_load(*args):
    global _TEXT_OUTPUT
    doc = CalcDoc.from_current_doc()
    sheet = doc.sheets[0]

    cell = sheet[1, 1]
    ctl = cell.control.insert_control_text_field()
    ctl.model.BackgroundColor = StandardColor.GREEN_LIGHT3
    ctl.add_event_mouse_pressed(on_mouse_event)
    ctl.add_event_focus_lost(on_lost_focus)
    rng = sheet.get_range(range_name="F2:K10")
    ctl_text = rng.control.insert_control_text_field()
    ctl_text.model.BackgroundColor = StandardColor.BLUE_LIGHT3
    ctl_text.set_property(
        VerticalAlign=VerticalAlignment.TOP, VScroll=True, ReadOnly=True, MultiLine=True
    )
    _TEXT_OUTPUT = ctl_text
    doc.toggle_design_mode()


g_exportedScripts = (on_load,)
