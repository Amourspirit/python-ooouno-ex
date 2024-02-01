# region Imports
from __future__ import annotations
import uno
from typing import Any, cast, TYPE_CHECKING, Tuple

from ooo.dyn.style.vertical_alignment import VerticalAlignment

from ooodev.dialog import Dialogs, BorderKind, TriStateKind
from ooodev.dialog.input import Input
from ooodev.events.args.event_args import EventArgs
from ooodev.dialog import Dialog

if TYPE_CHECKING:
    from com.sun.star.awt import ActionEvent

    # from com.sun.star.awt import ItemEvent
    from com.sun.star.awt import KeyEvent
    from com.sun.star.awt import XControl
    from com.sun.star.awt.tree import MutableTreeNode
    from com.sun.star.awt.tree import TreeDataModelEvent
    from com.sun.star.awt.tree import TreeExpansionEvent
    from com.sun.star.sheet import XSpreadsheetDocument
    from ooodev.adapter.tree.tree_data_model_comp import TreeDataModelComp
    from ooodev.dialog.dl_control.ctl_tree import CtlTree
    from ooodev.dialog.dl_control.ctl_check_box import CtlCheckBox


# endregion Imports


class TreeSimple:
    """Tree simple example."""

    # pylint: disable=unused-argument
    # region Init
    def __init__(
        self,
        dialog: Dialog,
        ctrl: XControl,
        x: int,
        y: int,
        width: int,
        height: int,
        border_kind: BorderKind,
    ) -> None:
        self._dialog = dialog
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
        self._box_height = 20
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
        self._init_options()

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

        self._fn_on_action_preformed_btn = self.on_action_preformed_btn
        self._fn_on_focus_gained = self.on_focus_gained
        self._fn_on_focus_lost = self.on_focus_lost
        self._fn_on_option_edit_state_changed = self.on_option_edit_state_changed
        self._fn_on_tree_key_released = self.on_tree_key_released
        self._fn_on_tree_node_collapsed = self.on_tree_node_collapsed
        self._fn_on_tree_node_collapsing = self.on_tree_node_collapsing
        self._fn_on_tree_node_expanded = self.on_tree_node_expanded
        self._fn_on_tree_node_expanding = self.on_tree_node_expanding
        self._fn_on_tree_nodes_changed = self.on_tree_nodes_changed
        self._fn_on_tree_nodes_inserted = self.on_tree_nodes_inserted
        self._fn_on_tree_nodes_removed = self.on_tree_nodes_removed
        self._fn_on_tree_request_child_nodes = self.on_tree_request_child_nodes
        self._fn_on_tree_selection_changed = self.on_tree_selection_changed
        self._fn_on_tree_tree_structure_changed = self.on_tree_tree_structure_changed

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = self._dialog.insert_label(
            label=self.get_label_msg(),
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=self._box_height * 2,
            MultiLine=True,
            dialog_ctrl=self._control,
        )

    def _init_tree(self) -> None:
        sz_main_lbl = self._ctl_main_lbl.view.getPosSize()
        # multi_select must be false for drop_down to work.
        self._ctl_tree = self._dialog.insert_tree_control(
            x=sz_main_lbl.X,
            y=sz_main_lbl.Y + sz_main_lbl.Height + self._vert_margin,
            width=round(sz_main_lbl.Width / 2) - self._margin,
            # height=self._height - sz.Height - self._vert_margin,
            height=round(
                self._height
                - sz_main_lbl.Height
                - self._btn_height
                - (self._vert_margin * 2)
            ),
            border=self._border_kind,
            dialog_ctrl=self._control,
        )
        self._ctl_tree.model.InvokesStopNodeEditing = True
        self._ctl_tree.model.Editable = True
        self._init_tree_events()
        self._data_model = self._ctl_tree.data_model
        self._init_data_model_events()

    def _init_data_model_events(self) -> None:
        if self._data_model is not None:
            self._data_model.add_event_tree_nodes_changed(
                self._fn_on_tree_nodes_changed
            )
            self._data_model.add_event_tree_nodes_inserted(
                self._fn_on_tree_nodes_inserted
            )
            self._data_model.add_event_tree_nodes_removed(
                self._fn_on_tree_nodes_removed
            )
            self._data_model.add_event_tree_structure_changed(
                self._fn_on_tree_tree_structure_changed
            )

    def _init_event_text(self) -> None:
        sz_tree = self._ctl_tree.view.getPosSize()
        sz_main_lbl = self._ctl_main_lbl.view.getPosSize()
        self._event_text = self._dialog.insert_text_field(
            text="",
            x=sz_tree.X + sz_tree.Width + (self._margin * 2),
            y=sz_tree.Y,
            width=sz_tree.Width,
            # height=sz.Height - self._btn_height - (self._vert_margin * 2),
            height=round(
                self._height
                - sz_main_lbl.Height
                - self._btn_height
                - (self._vert_margin * 2)
            ),
            border=self._border_kind,
            dialog_ctrl=self._control,
            VerticalAlign=VerticalAlignment.TOP,
            ReadOnly=True,
            MultiLine=True,
            AutoVScroll=True,
        )

    def _init_buttons(self) -> None:
        """Add OK, Cancel and Info buttons to dialog control"""
        sz = self._event_text.view.getPosSize()
        self._ctl_btn_clear = self._dialog.insert_button(
            label="Clear",
            x=sz.X + sz.Width - self._btn_width,
            y=sz.Y + sz.Height + self._vert_margin,
            width=self._btn_width,
            height=self._btn_height,
            dialog_ctrl=self._control,
        )
        self._ctl_btn_clear.view.setActionCommand("CLEAR")
        self._ctl_btn_clear.model.HelpText = "Clear contents"
        self._ctl_btn_clear.add_event_action_performed(self._fn_on_action_preformed_btn)

        sz_btn = self._ctl_btn_clear.view.getPosSize()
        self._ctl_btn_search = self._dialog.insert_button(
            label="Search...",
            x=sz_btn.X - sz_btn.Width - self._padding,
            y=sz_btn.Y,
            width=self._btn_width,
            height=self._btn_height,
            dialog_ctrl=self._control,
        )
        self._ctl_btn_search.view.setActionCommand("SEARCH")
        self._ctl_btn_search.model.HelpText = "Search the Tree nodes"
        self._ctl_btn_search.add_event_action_performed(
            self._fn_on_action_preformed_btn
        )

    def _init_options(self) -> None:
        """Add OK, Cancel and Info buttons to dialog control"""
        sz_tree = self._ctl_tree.view.getPosSize()
        self._ctl_chk_node_edit = self._dialog.insert_check_box(
            label="Allow Tree Node Editing",
            x=sz_tree.X,
            y=sz_tree.Y + sz_tree.Height + self._vert_margin,
            width=200,
            height=self._box_height,
            tri_state=False,
            dialog_ctrl=self._control,
        )
        self._ctl_chk_node_edit.tip_text = "Specifies if the tree nodes can be edited."
        self._ctl_chk_node_edit.add_event_item_state_changed(
            self._fn_on_option_edit_state_changed
        )
        self._ctl_chk_node_edit.state = TriStateKind.CHECKED

    def _init_tree_events(self) -> None:
        self.control_tree.add_event_selection_changed(
            self._fn_on_tree_selection_changed
        )
        self.control_tree.add_event_tree_collapsing(self._fn_on_tree_node_collapsing)
        self.control_tree.add_event_tree_expanding(self._fn_on_tree_node_expanding)
        self.control_tree.add_event_tree_expanded(self._fn_on_tree_node_expanded)
        self.control_tree.add_event_tree_collapsed(self._fn_on_tree_node_collapsed)
        self.control_tree.add_event_request_child_nodes(
            self._fn_on_tree_request_child_nodes
        )
        self.control_tree.add_event_key_released(self._fn_on_tree_key_released)
        self.control_tree.add_event_focus_gained(self._fn_on_focus_gained)
        self.control_tree.add_event_focus_lost(self._fn_on_focus_lost)

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
            sub_node = self._ctl_tree.add_sub_node(
                parent_node=self._root_node, display_value=f"Node {i + 1}"
            )
            for i in range(8):
                sub_sub_node = self._ctl_tree.add_sub_node(
                    parent_node=sub_node, display_value=f"Sub Node {i + 1}"
                )
                for i in range(3):
                    _ = self._ctl_tree.add_sub_node(
                        parent_node=sub_sub_node, display_value=f"Sub Sub Node {i + 1}"
                    )

    # endregion Data

    # region Search
    def _search_nodes(self) -> None:
        """Search the tree nodes for the search text."""
        search_text = ""
        if self._selected_node is not None:
            search_text = Input.get_input(
                "Search Text", "Enter the text to search for:", search_text
            )
            if search_text:
                self._event_text.write_line(
                    f"Searching for '{search_text}' in '{self._selected_node.getDisplayValue()}'"
                )
                node = cast(
                    "MutableTreeNode",
                    self._ctl_tree.find_node(
                        node=self._selected_node,
                        value=search_text,
                        case_sensitive=False,
                        search_data_value=True,
                    ),
                )
                if node:
                    parent = cast("MutableTreeNode", node.getParent())
                    while parent:
                        self._ctl_tree.view.expandNode(parent)
                        parent = cast("MutableTreeNode", parent.getParent())
                    # self._ctl_tree.view.expandNode(node)
                    self._ctl_tree.view.clearSelection()
                    self._ctl_tree.view.addSelection(node)
                    self._event_text.write_line(f"Found: '{node.getDisplayValue()}'")
                    if node.DataValue:
                        self._event_text.write_line(f"Data Value: '{node.DataValue}'")
                else:
                    self._event_text.write_line(f"Node Not Found: '{search_text}'")

    # endregion Search

    # region Event Handlers
    def on_option_edit_state_changed(
        self, src: Any, event: EventArgs, control_src: CtlCheckBox, *args, **kwargs
    ) -> None:
        if control_src.state == TriStateKind.CHECKED:
            self._ctl_tree.model.Editable = True
        else:
            self._ctl_tree.model.Editable = False

        self._event_text.write_line(
            f"Tree Control Editable: {self._ctl_tree.model.Editable}"
        )

    def on_tree_key_released(
        self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs
    ) -> None:
        itm_event = cast("KeyEvent", event.event_data)
        self._event_text.write_line(f"Key Released KeyCode: {itm_event.KeyCode}")
        self._event_text.write_line(f"Key Released Modifiers: {itm_event.Modifiers}")

    def on_focus_gained(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        self._event_text.write_line(f"Focus Gained: {control_src.name}")

    def on_focus_lost(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        self._event_text.write_line(f"Focus Lost: {control_src.name}")

    def on_tree_selection_changed(
        self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs
    ) -> None:
        self._selected_node = control_src.current_selection
        if self._selected_node is not None:
            self._event_text.write_line(
                f"Selection changed: {self._selected_node.getDisplayValue()}"
            )
            if self._selected_node.DataValue is not None:
                self._event_text.write_line(
                    f"Node Data Value: {self._selected_node.DataValue}"
                )
        else:
            self._event_text.write_line("Selection changed: None")

    def on_tree_node_collapsing(
        self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs
    ) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(
            f"Node Collapsing: {itm_event.Node.getDisplayValue()}"
        )

    def on_tree_node_collapsed(
        self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs
    ) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(
            f"Node Collapsed: {itm_event.Node.getDisplayValue()}"
        )

    def on_tree_node_expanding(
        self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs
    ) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(
            f"Node Expanding: {itm_event.Node.getDisplayValue()}"
        )

    def on_tree_node_expanded(
        self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs
    ) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(
            f"Node Expanded: {itm_event.Node.getDisplayValue()}"
        )

    def on_tree_request_child_nodes(
        self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs
    ) -> None:
        itm_event = cast("TreeExpansionEvent", event.event_data)
        self._event_text.write_line(
            f"Node Request Child Nodes: {itm_event.Node.getDisplayValue()}"
        )

    def on_tree_nodes_changed(
        self,
        src: Any,
        event: EventArgs,
        control_src: TreeDataModelComp,
        *args,
        **kwargs,
    ) -> None:
        itm_event = cast("TreeDataModelEvent", event.event_data)
        parent_name = (
            itm_event.ParentNode.getDisplayValue()
            if itm_event.ParentNode is not None
            else ""
        )
        for node in itm_event.Nodes:
            if parent_name:
                self._event_text.write_line(
                    f"Node Changed: {parent_name} -> {node.getDisplayValue()}"
                )
            else:
                self._event_text.write_line(f"Node Changed: {node.getDisplayValue()}")

    def on_tree_nodes_inserted(
        self,
        src: Any,
        event: EventArgs,
        control_src: TreeDataModelComp,
        *args,
        **kwargs,
    ) -> None:
        itm_event = cast("TreeDataModelEvent", event.event_data)
        parent_name = (
            itm_event.ParentNode.getDisplayValue()
            if itm_event.ParentNode is not None
            else ""
        )
        for node in itm_event.Nodes:
            if parent_name:
                self._event_text.write_line(
                    f"Node Inserted: {parent_name} -> {node.getDisplayValue()}"
                )
            else:
                self._event_text.write_line(f"Node Inserted: {node.getDisplayValue()}")

    def on_tree_nodes_removed(
        self,
        src: Any,
        event: EventArgs,
        control_src: TreeDataModelComp,
        *args,
        **kwargs,
    ) -> None:
        itm_event = cast("TreeDataModelEvent", event.event_data)
        parent_name = (
            itm_event.ParentNode.getDisplayValue()
            if itm_event.ParentNode is not None
            else ""
        )
        for node in itm_event.Nodes:
            if parent_name:
                self._event_text.write_line(
                    f"Node Removed: {parent_name} -> {node.getDisplayValue()}"
                )
            else:
                self._event_text.write_line(f"Node Removed: {node.getDisplayValue()}")

    def on_tree_tree_structure_changed(
        self,
        src: Any,
        event: EventArgs,
        control_src: TreeDataModelComp,
        *args,
        **kwargs,
    ) -> None:
        itm_event = cast("TreeDataModelEvent", event.event_data)
        parent_name = (
            itm_event.ParentNode.getDisplayValue()
            if itm_event.ParentNode is not None
            else ""
        )
        for node in itm_event.Nodes:
            if parent_name:
                self._event_text.write_line(
                    f"Tree Structure Changed: {parent_name} -> {node.getDisplayValue()}"
                )
            else:
                self._event_text.write_line(
                    f"Tree Structure Changed: {node.getDisplayValue()}"
                )

    def on_action_preformed_btn(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        itm_event = cast("ActionEvent", event.event_data)
        if itm_event.ActionCommand == "CLEAR":
            self._event_text.text = ""
        elif itm_event.ActionCommand == "SEARCH":
            self._search_nodes()

    # endregion Event Handlers

    # region Properties
    @property
    def dialog(self) -> Dialog:
        return self._dialog
    
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
