# region Imports
from __future__ import annotations
import contextlib
import os
import uno
from typing import Any, cast, TYPE_CHECKING, Tuple

from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.selection import Selection

from ooodev.dialog import Dialogs, BorderKind
from ooodev.events.args.event_args import EventArgs
from ooodev.office.calc import Calc

if TYPE_CHECKING:
    from com.sun.star.awt import ActionEvent
    from com.sun.star.awt import ItemEvent
    from com.sun.star.awt.tree import MutableTreeNode
    from com.sun.star.sheet import XSpreadsheetDocument
    from com.sun.star.awt import XControl
    from ooodev.dialog.dl_control.ctl_tree import CtlTree
    from com.sun.star.awt.tree import TreeExpansionEvent


# endregion Imports


class TreeSimple:
    """Tree simple example."""

    # pylint: disable=unused-argument
    # region Init
    def __init__(
        self,
        ctrl: XControl,
        x: int,
        y: int,
        width: int,
        height: int,
        border_kind: BorderKind,
    ) -> None:
        self._control = ctrl
        self._x = x
        self._y = y
        self._width = width
        self._height = height
        self._border_kind = border_kind
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 6
        self._vert_margin = 12
        self._box_height = 30
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._row_index = -1
        self._selected_node: MutableTreeNode | None = None
        self._init()

    def _init(self) -> None:
        self._init_handlers()
        self._init_label()
        self._init_tree()
        self._init_event_text()
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

        self._fn_on_tree_selection_changed = self.on_tree_selection_changed
        self._fn_on_tree_node_collapsing = self.on_tree_node_collapsing
        self._fn_on_tree_node_collapsed = self.on_tree_node_collapsed
        self._fn_on_tree_node_expanding = self.on_tree_node_expanding
        self._fn_on_tree_node_expanded = self.on_tree_node_expanded
        self._fn_on_tree_request_child_nodes = self.on_tree_request_child_nodes
        self._fn_on_action_preformed_btn = self.on_action_preformed_btn

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = Dialogs.insert_label(
            dialog_ctrl=self._control,
            label=self.get_label_msg(),
            x=self._margin,
            y=self._padding,
            width=self._width - (self._margin * 2),
            height=self._box_height,
        )

    def _init_tree(self) -> None:
        sz = self._ctl_main_lbl.view.getPosSize()
        # multi_select must be false for drop_down to work.
        self._ctl_tree = Dialogs.insert_tree_control(
            dialog_ctrl=self._control,
            x=sz.X,
            y=sz.Y + sz.Height + self._vert_margin,
            width=round(sz.Width / 2) - self._margin,
            height=self._height - sz.Height - self._vert_margin,
            border=self._border_kind,
        )

        self._init_tree_events()

    def _init_event_text(self) -> None:
        sz = self._ctl_tree.view.getPosSize()
        self._event_text = Dialogs.insert_text_field(
            dialog_ctrl=self._control,
            text="",
            x=sz.X + sz.Width + (self._margin * 2),
            y=sz.Y,
            width=sz.Width,
            height=sz.Height - self._btn_height - (self._vert_margin * 2),
            border=self._border_kind,
            VerticalAlign=VerticalAlignment.TOP,
            ReadOnly=True,
            MultiLine=True,
            AutoVScroll=True,
        )

    def _init_buttons(self) -> None:
        """Add OK, Cancel and Info buttons to dialog control"""
        sz = self._event_text.view.getPosSize()
        self._ctl_btn_clear = Dialogs.insert_button(
            dialog_ctrl=self._control,
            label="Clear",
            x=sz.X + sz.Width - self._btn_width,
            y=sz.Y + sz.Height + self._vert_margin,
            width=self._btn_width,
            height=self._btn_height,
        )
        self._ctl_btn_clear.view.setActionCommand("CLEAR")
        self._ctl_btn_clear.model.HelpText = "Clear contents"
        self._ctl_btn_clear.add_event_action_performed(self._fn_on_action_preformed_btn)

    def _init_tree_events(self) -> None:
        self.control_tree.add_event_selection_changed(self._fn_on_tree_selection_changed)
        self.control_tree.add_event_tree_collapsing(self._fn_on_tree_node_collapsing)
        self.control_tree.add_event_tree_expanding(self._fn_on_tree_node_expanding)
        self.control_tree.add_event_tree_expanded(self._fn_on_tree_node_expanded)
        self.control_tree.add_event_tree_collapsed(self._fn_on_tree_node_collapsed)
        self.control_tree.add_event_request_child_nodes(self._fn_on_tree_request_child_nodes)

    # endregion Init

    def get_label_msg(self) -> str:
        return "This is an example that shows how to use a tree control.\nData is generated and added using the 'add_sub_node()' method."

    # region Data
    def _get_root_node(self) -> "MutableTreeNode":
        return self._root_node

    def get_selected_items(self) -> Tuple[int, ...]:
        """Get the items that are to be selected in the listbox at startup."""
        return (0,)

    def set_data(self) -> None:
        """Set the data in the tree control."""
        self._root_node = self._ctl_tree.create_root(display_value="Root")
        self._selected_node = self._get_root_node()
        for i in range(5):
            sub_node = self._ctl_tree.add_sub_node(parent_node=self._root_node, display_value=f"Node {i + 1}")
            for i in range(8):
                sub_sub_node = self._ctl_tree.add_sub_node(parent_node=sub_node, display_value=f"Sub Node {i + 1}")
                for i in range(3):
                    _ = self._ctl_tree.add_sub_node(parent_node=sub_sub_node, display_value=f"Sub Sub Node {i + 1}")

    # endregion Data

    # region Event Handlers
    def on_tree_selection_changed(self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs) -> None:
        self._selected_node = cast("MutableTreeNode", control_src.view.getSelection())
        if self._selected_node is not None:
            self._event_text.write_line(f"Selection changed: {self._selected_node.getDisplayValue()}")
        else:
            self._event_text.write_line("Selection changed: None")

    def on_tree_node_collapsing(self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(f"Node Collapsing: {itm_event.Node.getDisplayValue()}")

    def on_tree_node_collapsed(self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(f"Node Collapsed: {itm_event.Node.getDisplayValue()}")

    def on_tree_node_expanding(self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(f"Node Expanding: {itm_event.Node.getDisplayValue()}")

    def on_tree_node_expanded(self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(f"Node Expanded: {itm_event.Node.getDisplayValue()}")

    def on_tree_request_child_nodes(self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(f"Node Request Child Nodes: {itm_event.Node.getDisplayValue()}")

    def on_action_preformed_btn(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        itm_event = cast("ActionEvent", event.event_data)
        if itm_event.ActionCommand == "CLEAR":
            self._event_text.text = ""

    # endregion Event Handlers

    # region Properties
    @property
    def x(self) -> int:
        return self._x

    @property
    def y(self) -> int:
        return self._y

    @property
    def width(self) -> int:
        return self._width

    @property
    def height(self) -> int:
        return self._height

    @property
    def border_kind(self) -> BorderKind:
        return self._border_kind

    @property
    def btn_width(self) -> int:
        return self._btn_width

    @property
    def btn_height(self) -> int:
        return self._btn_height

    @property
    def margin(self) -> int:
        return self._margin

    @property
    def vert_margin(self) -> int:
        return self._vert_margin

    @property
    def box_height(self) -> int:
        return self._box_height

    @property
    def padding(self) -> int:
        return self._padding

    @property
    def doc(self) -> XSpreadsheetDocument:
        return self._doc

    @property
    def control(self) -> XControl:
        return self._control

    @property
    def control_tree(self) -> CtlTree:
        return self._ctl_tree

    @property
    def selected_item(self) -> str:
        return self._selected_item

    @property
    def selected_node(self) -> MutableTreeNode | None:
        return self._selected_node

    @selected_node.setter
    def selected_node(self, value: MutableTreeNode | None) -> None:
        self._selected_node = value

    # endregion Properties
