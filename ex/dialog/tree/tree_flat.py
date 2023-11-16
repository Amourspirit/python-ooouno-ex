from __future__ import annotations
from tree_simple import TreeSimple


class TreeFlat(TreeSimple):
    """Tree control with a flat list of items."""

    def set_data(self) -> None:
        """Set the data in the tree control."""
        self._root_node = self.control_tree.create_root(display_value="Root")
        self.selected_node = self._get_root_node()
        # flat_tree_data = ["Node " + str(i) for i in range(1, 31)]
        flat_list = [
            ["A1", "B1", "C1"],
            ["A1", "B1", "C2"],
            ["A1", "B2", "C3"],
            ["A2", "B3", "C4"],
            ["A2", "B3", "C5"],
            ["A2", "B3", "C6"],
            ["A2", "B4", "Razor"],
        ]
        self.control_tree.add_sub_tree(flat_tree=flat_list, parent_node=self._root_node)

    def get_label_msg(self) -> str:
        return "This is an example that shows how to use a tree control.\nData is loaded from a flat list of items."
