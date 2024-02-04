# region Imports
from __future__ import annotations
import uno
from typing import Any, cast, TYPE_CHECKING
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.pos_size import PosSize
from ooodev.dialog.msgbox import MessageBoxResultsEnum, MessageBoxType

from ooodev.dialog import BorderKind
from ooodev.events.args.event_args import EventArgs
from ooodev.calc import CalcDoc

if TYPE_CHECKING:
    from com.sun.star.awt import ActionEvent
    from com.sun.star.awt.grid import GridSelectionEvent
    from ooodev.dialog.dl_control.ctl_grid import CtlGrid


# endregion Imports


class GridEx:
    # pylint: disable=unused-argument
    # region Init
    def __init__(self, doc: CalcDoc) -> None:
        self._border_kind = BorderKind.BORDER_SIMPLE
        self._width = 600
        self._height = 410
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 6
        self._vert_margin = 12
        self._box_height = 30
        self._title = "Grid Example"
        self._msg = "This is a grid example."
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._row_index = -1
        self._doc = doc
        self._init_dialog()
        self._sheet = doc.get_active_sheet()
        self._sheet["A1"].goto()
        self._set_table_data()

    def _init_dialog(self) -> None:
        """Create dialog and add controls."""
        self._init_handlers()
        self._dialog = self._doc.create_dialog(
            x=-1, y=-1, width=self._width, height=self._height, title=self._title
        )
        self._init_label()
        self._init_table()
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

        self._fn_grid_selection_changed = self.on_grid_selection_changed
        self._fn_button_action_preformed = self.on_button_action_preformed

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = self._dialog.insert_label(
            label=self._msg,
            x=self._margin,
            y=self._padding,
            width=self._width - (self._padding * 2),
            height=self._box_height,
        )

    def _init_table(self) -> None:
        """Add a Grid (table) to the dialog control"""
        sz = self._ctl_main_lbl.view.getPosSize()
        self._ctl_table1 = self._dialog.insert_table_control(
            x=sz.X,
            y=sz.Y + sz.Height + self._margin,
            width=sz.Width,
            height=300,
            grid_lines=True,
            col_header=True,
            row_header=True,
        )
        self._ctl_table1.add_event_selection_changed(self._fn_grid_selection_changed)

    def _init_buttons(self) -> None:
        """Add OK, Cancel and Info buttons to dialog control"""
        self._ctl_btn_cancel = self._dialog.insert_button(
            label="Cancel",
            x=self._width - self._btn_width - self._margin,
            y=self._height - self._btn_height - self._vert_margin,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.CANCEL,
        )
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

        self._ctl_btn_info = self._dialog.insert_button(
            label="Info",
            x=self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
        )
        self._ctl_btn_info.view.setActionCommand("INFO")
        self._ctl_btn_info.model.HelpText = "Show info for selected item."
        self._ctl_btn_info.add_event_action_performed(self._fn_button_action_preformed)

    # endregion Init

    # region Data
    def _set_table_data(self) -> None:
        """Find all the data in the spreadsheet and add it to the dialog grid control."""
        rng = self._sheet.find_used_range()
        tbl = rng.get_array()
        self._ctl_table1.set_table_data(
            data=tbl,
            align="RLC",
            # widths=widths,
            has_row_headers=False,
            has_colum_headers=True,
        )

    # endregion Data

    # region Handle Results
    def _handle_results(self, result: int) -> None:
        """Display a message if the OK button has been clicked"""
        if result == MessageBoxResultsEnum.OK.value:
            if self._row_index >= 0:
                # get the selected row from the grid. Returns a Tuple of Objects.
                row = self._ctl_table1.model.GridDataModel.getRowData(self._row_index)
                msg = f"{row}"
                self._doc.msgbox(
                    msg=msg, title="Selected Values", boxtype=MessageBoxType.INFOBOX
                )
            else:
                self._doc.msgbox(
                    msg="Nothing was selected",
                    title="No Selection",
                    boxtype=MessageBoxType.INFOBOX,
                )

    # endregion Handle Results

    # region Event Handlers
    def on_grid_selection_changed(
        self, src: Any, event: EventArgs, control_src: CtlGrid, *args, **kwargs
    ) -> None:
        """Method that is fired each time the selection changes in the grid."""
        print("Grid Selection Changed:", control_src.name)
        itm_event = cast("GridSelectionEvent", event.event_data)
        self._row_index = (
            itm_event.SelectedRowIndexes[0] if itm_event.SelectedRowIndexes else -1
        )
        print("Selected row indexes:", itm_event.SelectedRowIndexes)
        print("Selected row index:", self._row_index)

    def on_button_action_preformed(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        """Method that is fired when Info button is clicked."""
        itm_event = cast("ActionEvent", event.event_data)
        if itm_event.ActionCommand == "INFO":
            self._handle_results(MessageBoxResultsEnum.OK.value)

    # endregion Event Handlers

    # region Show Dialog
    def show(self) -> int:
        # window = Lo.get_frame().getContainerWindow()
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
