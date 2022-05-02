# coding: utf-8
'''
Selection framework is a set of methods for accessing and reading LibreOffice Writer document contents.
'''
from __future__ import annotations
from typing import TYPE_CHECKING, Union
from .ooo_util import get_xModel

if TYPE_CHECKING:
    from ooo.lo.view.x_selection_supplier import XSelectionSupplier
    from ooo.lo.text.x_text_view_cursor import XTextViewCursor
    from ooo.lo.text.x_text_cursor import XTextCursor
    from ooo.lo.text.x_text import XText
    from ooo.lo.text.text_range import TextRange
    from ooo.lo.text.x_word_cursor import XWordCursor
    from ooo.lo.text.x_text_range import XTextRange
    from ooo.lo.text.generic_text_document import GenericTextDocument
    from ooo.lo.text.x_text_range_compare import XTextRangeCompare
    from ooo.lo.container.x_index_access import XIndexAccess
    from ooo.lo.frame.dispatch_helper import DispatchHelper
    from ooo.lo.frame.x_dispatch_provider import XDispatchProvider
    from ooo.lo.frame.x_model import XModel


def select_view_by_cursors(**kwargs):
    """
    Selects the Text View ( visible selection ) for the given cursors

    Keyword Args:
        sel (Tuple[XTextRange, XTextRange], XTextRange): selection as tuple of left and right range or as text range.
        o_doc (GenericTextDocument, optional): current document (xModel). Defaults to current document.
        o_text (XText, optional): xText object used only when sel is a xTextRangeObject.
        require_selection (bool, optional): If ``True`` then a check is preformed to see if anything is selected;
            Otherwise, No check is done. Default ``True``

    Raises:
        TypeError: if ``sel`` is ``None``
        ValueError: if ``sel`` is passed in as ``tuple`` and length is not ``2``.
        ValueError: if ``sel`` is missing.
        Excpetion: If Error selecting view.
    """
    o_doc: 'GenericTextDocument' = kwargs.get('o_doc', None)
    if o_doc is None:
        o_doc = get_xModel()
    _sel_check = kwargs.get('require_selection', True)

    if _sel_check == True and is_anything_selected(o_doc=o_doc) == False:
        return None
    l_cursor: 'XTextCursor' = None
    r_cursor: 'XTextCursor' = None
    _sel: 'Union[tuple, XTextRange]' = kwargs.get('sel', None)
    if _sel is None:
        raise ValueError("select_view_by_cursors() 'sel' argument is required")
    if isinstance(_sel, tuple):
        if len(_sel) < 2:
            raise ValueError(
                "select_view_by_cursors() sel argument when passed as a tuple is expected to have two elements")
        l_cursor = _sel[0]
        r_cursor = _sel[1]
    else:
        x_text: 'Union[XText, None]' = kwargs.get("o_text", None)
        if x_text is None:
            x_text = get_selected_text(o_doc=o_doc)
 
        if x_text == None:
            # there is an issue. Something should be selected.
            # msg = "select_view_by_cursors() Something was expected to be selected but xText object does not exist"
            return None
        l_cursor = _get_left_cursor(o_sel=_sel, o_text=x_text)
        r_cursor = _get_right_cursor(o_sel=_sel, o_text=x_text)

    vc = get_view_cursor(o_doc=o_doc)
    try:
        vc.setVisible(False)
        vc.gotoStart(False)
        vc.collapseToStart()
        vc.gotoRange(l_cursor, False)
        vc.gotoRange(r_cursor, True)
    except Exception as e:
        raise e
    finally:
        if not vc.isVisible():
            vc.setVisible(True)

def get_view_cursor(**kwargs) -> 'XTextViewCursor':
    """
    Gets current view cursor which is a XTextViewCursor

    Keyword Args:
        o_doc (object, optional): current document (xModel)

    Returns:
        object: View Cursor
    """
    o_doc = kwargs.get('o_doc', None)
    if o_doc is None:
        o_doc = get_xModel()
    # https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextViewCursor.html
    frame: object = o_doc.CurrentController.Frame
    current_controler: object = frame.getController()  # XController
    view_cursor = current_controler.getViewCursor()
    return view_cursor

def is_anything_selected(**kwargs) -> bool:
    """
    Determine if anything is selected.

    Keyword Args:
        o_doc (object, optional): current document (xModel)

    Returns:
        bool: ``True`` if anything in the document is selected: Otherwise, ``False``
    """
    if 'o_doc' in kwargs.keys():
        o_doc = kwargs['o_doc']
    else:
        o_doc = get_xModel()
    # print("is_anything_selected Method")

    o_selections = o_doc.getCurrentSelection()
    if not o_selections:
        # print("o_selections was not created from o_doc")
        return False

    count = int(o_selections.getCount())
    # print("Selections Count:", count)
    if count == 0:
        return False
    elif count > 1:
        return True
    else:
        # There is only one selection so obtain the first selection
        o_sel = o_selections.getByIndex(0)
        o_text = o_sel.getText()
        # Create a text cursor that covers the range and then see if it is collapsed
        o_cursor = o_text.createTextCursorByRange(o_sel)  # XTextCursor
        # print("o_cursor",o_cursor)
        if not o_cursor.isCollapsed():
            # print("o_cursor is NOT Collapsed")
            return True

    return False

def get_selected_text(**kwargs) -> 'Union[XText, None]':
    """
    Gets the xText for current selection

    Keyword Arguments:
        o_doc (object, optional): current document (xModel): Default: Current xModel

    Returns:
        Union[object, None]: If no selection is made then None is returned; Otherwise, xText is returned.
    """
    o_doc: 'Union[XModel, None]' = kwargs.get('o_doc', None)
    if o_doc is None:
        o_doc = get_xModel()
    o_selections: 'XIndexAccess' = o_doc.getCurrentSelection()
    if not o_selections:
        return None
    count = int(o_selections.getCount())
    if count == 0:
        return None
    o_sel: 'XText' = o_selections.getByIndex(0)
    return o_sel.getText()


def _get_right_cursor(o_sel: 'XTextRange', o_text: 'XText') -> 'XTextCursor':
    if o_text.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
        o_range = o_sel.getStart()
    else:
        o_range = o_sel.getEnd()
    cursor = o_text.createTextCursorByRange(o_range)
    cursor.goLeft(0, False)
    return cursor

def _get_left_cursor(o_sel: 'XTextRange', o_text: 'XText') -> 'XTextCursor':
    if o_text.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
        o_range = o_sel.getEnd()
    else:
        o_range = o_sel.getStart()
    cursor = o_text.createTextCursorByRange(o_range)
    cursor.goRight(0, False)
    return cursor