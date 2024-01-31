# region Imports
from __future__ import annotations
import uno
from typing import Any, cast, TYPE_CHECKING
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.pos_size import PosSize
from ooodev.dialog.msgbox import MessageBoxResultsEnum, MessageBoxType

from ooodev.dialog import BorderKind
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.color import StandardColor
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.awt import AdjustmentEvent
    from ooodev.dialog.dl_control.ctl_scroll_bar import CtlScrollBar
    from ooodev.dialog.dl_control.ctl_base import DialogControlBase

# endregion Imports


class Progress:
    # pylint: disable=unused-argument
    # region Init
    def __init__(self) -> None:
        self._doc = Lo.current_doc
        self._border_kind = BorderKind.BORDER_3D
        self._width = 500
        self._height = 180
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 6
        self._box_height = 30
        self._title = "Progress Scroll"
        self._msg = "Progress and Scrollbar example."
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._progress_value = 67
        self._current_tab_index = 1
        self._min_value = 1
        self._max_value = 100
        self._init_dialog()

    def _init_dialog(self) -> None:
        """Create dialog and add controls."""
        self._init_handlers()
        self._dialog = self._doc.create_dialog(
            x=-1, y=-1, width=self._width, height=self._height, title=self._title
        )
        self._dialog.create_peer()
        self._init_label()
        self._init_progress()
        self._init_scroll()
        self._init_buttons()

    def _init_handlers(self) -> None:
        """
        Add event handlers for when changes occur.

        Methods can not be assigned directly to control callbacks.
        This is a python thing. However, methods can be assigned to class
        variable an in turn those can be assigned to callbacks.

        Example:
            ``self._ctl_btn_info.add_event_action_performed(self.on_button_action_preformed)``
            This would not work!

            ``self._ctl_btn_info.add_event_action_performed(self._fn_button_action_preformed)``
            This will work.
        """

        self._fn_on_mouse_entered = self.on_mouse_entered
        self._fn_on_mouse_exit = self.on_mouse_exit
        self._fn_on_scroll_adjustment = self.on_scroll_adjustment

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = self._dialog.insert_label(
            label=self._msg,
            x=self._margin,
            y=self._padding,
            width=self._width - (self._margin * 2),
            height=self._box_height,
        )

    def _init_progress(self) -> None:
        sz = self._ctl_main_lbl.view.getPosSize()
        self._ctl_progress = self._dialog.insert_progress_bar(
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=sz.Width,
            height=self._box_height,
            min_value=self._min_value,
            max_value=self._max_value,
            value=self._progress_value,
            border=self._border_kind,
        )
        self._set_tab_index(self._ctl_progress)
        self._ctl_progress.fill_color = StandardColor.GREEN
        self._ctl_progress.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_progress.add_event_mouse_exited(self._fn_on_mouse_exit)

    def _init_scroll(self) -> None:
        sz = self._ctl_progress.view.getPosSize()
        self._ctl_scroll_progress = self._dialog.insert_scroll_bar(
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=sz.Width,
            height=14,
            min_value=self._ctl_progress.model.ProgressValueMin,
            max_value=self._ctl_progress.model.ProgressValueMax,
            border=self._border_kind,
        )
        self._set_tab_index(self._ctl_scroll_progress)
        self._ctl_scroll_progress.value = self._progress_value
        self._ctl_scroll_progress.add_event_adjustment_value_changed(
            self._fn_on_scroll_adjustment
        )

    def _init_buttons(self) -> None:
        """Add OK, Cancel and Info buttons to dialog control"""
        self._ctl_btn_cancel = self._dialog.insert_button(
            label="Cancel",
            x=self._width - self._btn_width - self._margin,
            y=self._height - self._btn_height - self._padding,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.CANCEL,
        )
        self._set_tab_index(self._ctl_btn_cancel)
        sz = self._ctl_btn_cancel.view.getPosSize()
        self._ctl_btn_ok = self._dialog.insert_button(
            label="OK",
            x=sz.X - sz.Width - self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )
        self._set_tab_index(self._ctl_btn_ok)

    # endregion Init

    # region Handle Results
    def _handle_results(self, result: int) -> None:
        """Display a message if the OK button has been clicked"""
        if result == MessageBoxResultsEnum.OK.value:
            _ = self._doc.msgbox(
                msg=f"Current Value: {self._progress_value}",
                title="Selected Values",
                boxtype=MessageBoxType.INFOBOX,
            )

    # endregion Handle Results

    # region Event Handlers
    def on_scroll_adjustment(
        self, src: Any, event: EventArgs, control_src: CtlScrollBar, *args, **kwargs
    ) -> None:
        # print("Scroll:", control_src.name)
        a_event = cast("AdjustmentEvent", event.event_data)
        self._progress_value = a_event.Value
        self._ctl_progress.value = self._progress_value

    def on_mouse_entered(
        self,
        src: Any,
        event: EventArgs,
        control_src: DialogControlBase,
        *args,
        **kwargs,
    ) -> None:
        # print(control_src)
        print("Mouse Entered:", control_src.name)

    def on_mouse_exit(
        self,
        src: Any,
        event: EventArgs,
        control_src: DialogControlBase,
        *args,
        **kwargs,
    ) -> None:
        # print(control_src)
        print("Mouse Exited:", control_src.name)

    # endregion Event Handlers

    # region Show Dialog
    def show(self) -> int:
        # make sure the document is active.
        self._doc.activate()
        window = self._doc.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.set_pos_size(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.set_visible(True)
        result = self._dialog.execute()
        self._handle_results(result)
        self._dialog.dispose()
        return result

    # endregion Show Dialog

    def _set_tab_index(self, ctl: DialogControlBase) -> None:
        ctl.tab_index = self._current_tab_index
        self._current_tab_index += 1
