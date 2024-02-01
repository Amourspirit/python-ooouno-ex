from __future__ import annotations
from typing import cast, TYPE_CHECKING
import re
from ooodev.utils.lo import Lo
from ooodev.dialog.search.tree_search import SearchTree, RuleTextRegex, RuleDataRegex

from tree_simple import TreeSimple

if TYPE_CHECKING:
    from com.sun.star.awt.tree import MutableTreeNode

# This class is used to demonstrate searching a tree control using Regular Expressions.
# OooDev has a SearchTree class that can be used to search a tree control.
# The SearchTree class uses rules to determine how to search the tree control.
# The rules are applied in the order they are registered.
# The first rule that matches is used to determine if the node is a match.
# Custom rules can be created by implementing the RuleT Protocol. See https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/search/tree_search/rule_proto.html
# See: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/search/tree_search/index.html
# for more information and rules for searching a tree control.


class TreeSearchRe(TreeSimple):
    """Tree control with a flat list of items."""

    def get_label_msg(self) -> str:
        return "This is an example that shows search a tree control using Regular Expressions."

    def set_data(self) -> None:
        """Set the data in the tree control."""
        self._root_node = self.control_tree.create_root(display_value="Root")
        self.selected_node = self._get_root_node()
        # flat_tree_data = ["Node " + str(i) for i in range(1, 31)]
        flat_list = [
            [["Awesomeness", 1], ["Best 50", "Best one"], ["Clever People", 55]],
            [
                ["Awesomeness"],
                ["Best 50"],
                ["Constant Peek", "See me too"],
            ],
            [
                ["Awesomeness"],
                ["B2", "To be or not to be"],
                ["C3", "CP30"],
            ],
            [
                ["Always Present", "Twice as cool"],
                ["Slippery silk", None],
                ["C4", "May be explosive"],
            ],
            [
                ["Always Present"],
                ["Slippery silk"],
                ["Fast Thinking", "May be more explosive"],
            ],
            [
                ["Always Present"],
                ["Slippery silk"],
                ["Sunset View", "May be most explosive"],
            ],
            [
                ["Always Present"],
                ["Start somewhere", "In the beginning"],
                ["Razors edge", "On the cutting edge"],
            ],
        ]
        self.control_tree.add_sub_tree(flat_tree=flat_list, parent_node=self._root_node)

    # region Search
    def _search_nodes(self) -> None:
        """Search the tree nodes for the search text."""
        search_text = ".*P3\d"
        if self._selected_node is not None:
            search_text = Lo.current_doc.input_box(
                title="Regular Expression Search",
                msg="Enter the regular expression to search for:",
                input_value=search_text,
            )
            if search_text:
                try:
                    regex = re.compile(search_text)
                except Exception:
                    self._event_text.write_line(
                        f"Invalid Regular Expression: '{search_text}'"
                    )
                    return
                search = SearchTree(None)
                # by adding RuleTextRegex first it means that the search will be done on the display value first and then the data value.
                search.register_rule(RuleTextRegex(regex))
                search.register_rule(RuleDataRegex(regex))
                self._event_text.write_line(
                    f"Searching for '{search_text}' in '{self._selected_node.getDisplayValue()}'"
                )
                node = cast("MutableTreeNode", search.find_node(self._selected_node))
                if node:
                    parent = cast("MutableTreeNode", node.getParent())
                    while parent:
                        self._ctl_tree.view.expandNode(parent)
                        parent = cast("MutableTreeNode", parent.getParent())
                    self._ctl_tree.view.clearSelection()
                    self._ctl_tree.view.addSelection(node)
                    self._event_text.write_line(f"Found: '{node.getDisplayValue()}'")
                    if node.DataValue:
                        self._event_text.write_line(f"Data Value: '{node.DataValue}'")
                else:
                    self._event_text.write_line(f"Node Not Found: '{search_text}'")

    # endregion Search
