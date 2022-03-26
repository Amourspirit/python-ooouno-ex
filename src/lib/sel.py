# coding: utf-8
"""
Selection framework is a set of methods for accessing and reading LibreOffice Writer document contents.
"""
from ooo.lo.beans.property_value import PropertyValue
from ooo.lo.container.x_index_access import XIndexAccess
from ooo.lo.frame.x_model import XModel
from ooo.lo.frame.dispatch_helper import DispatchHelper
from ooo.lo.frame.x_dispatch_provider import XDispatchProvider
from ooo.lo.frame.x_frame import XFrame
from ooo.lo.text.generic_text_document import GenericTextDocument
from ooo.lo.text.x_text import XText
from ooo.lo.view.x_selection_supplier import XSelectionSupplier
from ooo.lo.text.x_text_cursor import XTextCursor
from ooo.lo.text.x_text_range import XTextRange
from ooo.lo.text.x_text_range_compare import XTextRangeCompare
from ooo.lo.text.x_text_view_cursor import XTextViewCursor
from ooo.lo.text.x_text_view_cursor_supplier import XTextViewCursorSupplier
from ooo.lo.text.x_word_cursor import XWordCursor

from typing import Optional, Tuple, Union

from . import enums
from . import ooo_util


# region count


def compare_cursor_starts(
    o_text: XTextRangeCompare, c1: XTextRange, c2: XTextRange
) -> enums.CompareEnum:
    """
    Compares two cursors ranges start position

    Args:
        o_text (XTextRangeCompare): usually document text object
        c1 (XTextRange): first cursor range
        c2 (XTextRange): second cursor range

    Raises:
        Exception: if comparsion fails

    Returns:
        CompareEnum: | ``CompareEnum.BEFORE`` if :paramref:`~.compare_cursor_starts.c1` start position is before :paramref:`~.compare_cursor_starts.c2` start position
        | ``CompareEnum.EQUAL`` if :paramref:`~.compare_cursor_starts.c1` start position is equal to :paramref:`~.compare_cursor_starts.c2` start position
        | ``CompareEnum.AFTER`` if :paramref:`~.compare_cursor_starts.c1` start position is after :paramref:`~.compare_cursor_starts.c2` start position
    """
    i = o_text.compareRegionStarts(c1, c2)
    if i == 1:
        return enums.CompareEnum.BEFORE
    if i == -1:
        return enums.CompareEnum.AFTER
    if i == 0:
        return enums.CompareEnum.EQUAL
    # if no valid result raise error
    msg = "get_cursor_compare_starts() unable to get a valid compare result"
    raise Exception(msg)


def compare_cursor_ends(
    o_text: XTextRangeCompare, c1: XTextRange, c2: XTextRange
) -> enums.CompareEnum:
    """
    Compares two cursors ranges end positons

    Args:
        o_text (XTextRangeCompare): usually document text object
        c1 (XTextRange): first cursor range
        c2 (XTextRange): second cursor range

    Raises:
        Exception: if comparsion fails

    Returns:
        CompareEnum: | ``CompareEnum.BEFORE`` if :paramref:`~.compare_cursor_ends.c1` end position is before :paramref:`~.compare_cursor_ends.c2` end position
        | ``CompareEnum.EQUAL`` if :paramref:`~.compare_cursor_ends.c1` end position is equal to :paramref:`~.compare_cursor_ends.c2` end position
        | ``CompareEnum.AFTER`` if :paramref:`~.compare_cursor_ends.c1` end position is after :paramref:`~.compare_cursor_ends.c2` end position
    """

    i = o_text.compareRegionEnds(c1, c2)
    if i == 1:
        return enums.CompareEnum.BEFORE
    if i == -1:
        return enums.CompareEnum.AFTER
    if i == 0:
        return enums.CompareEnum.EQUAL
    # if no valid result raise error
    msg = "get_cursor_compare_ends() unable to get a valid compare result"
    raise Exception(msg)


def ooo_range_len(o_sel: XTextCursor, o_text: XText) -> int:
    """
    Gets the distance between range start and range end.

    Args:
        o_sel (XTextCursor): first cursor range
        o_text (XText): xText object, usually document text object

    Returns:
        int: length of range

    Note:
        All characters are counted including paragraph breaks.
        In Writer it will display selected characters however,
        paragraph breaks are not counted.
    """
    i = 0
    if o_sel.isCollapsed():
        return i
    l_cursor = _get_left_cursor(o_sel=o_sel, o_text=o_text)
    r_cursor = _get_right_cursor(o_sel=o_sel, o_text=o_text)
    if (
        compare_cursor_ends(o_text=o_text, c1=l_cursor, c2=r_cursor)
        < enums.CompareEnum.EQUAL
    ):
        while (
            compare_cursor_ends(o_text=o_text, c1=l_cursor, c2=r_cursor)
            != enums.CompareEnum.EQUAL
        ):
            l_cursor.goRight(1, False)
            i += 1
    return i


def ooo_range_distance(
    start_cursor: XTextCursor, end_cursor: XTextCursor, o_text: XText
) -> int:
    """
    Gets the distance between l_cursor start and r_cursor end.

    Args:
        start_cursor (XTextCursor): first cursor range
        end_cursor (XTextCursor): second cursor range
        o_text (XText): xText object, usually document text object

    Returns:
        int: length of range

    Note:
        All characters are counted including paragraph breaks.
        In Writer it will display selected characters however,
        paragraph breaks are not counted.
    """
    l_cursor = _get_left_cursor(o_sel=start_cursor, o_text=o_text)
    r_cursor = _get_right_cursor(o_sel=end_cursor, o_text=o_text)
    i = 0
    if (
        compare_cursor_ends(o_text=o_text, c1=l_cursor, c2=r_cursor)
        < enums.CompareEnum.EQUAL
    ):
        while (
            compare_cursor_ends(o_text=o_text, c1=l_cursor, c2=r_cursor)
            != enums.CompareEnum.EQUAL
        ):
            l_cursor.goRight(1, False)
            i += 1
    return i


# endregion count


def create_range_with_left_right(
    l_cursor: XTextCursor,
    r_cursor: XTextCursor,
    expand: bool,
    o_text: Optional[XText] = None,
) -> XTextCursor:
    """
    Combines left and right cursor into to one TextCursor.
    The return TextCursor can be read fom left to right

    Args:
        l_cursor (XTextCursor): left cursor Object
        r_cursor (XTextCursor): right cursor Object
        expand (bool): If ``True`` the return TextCursor will have range selected; Othwerwise, range will not be selected.
        o_text ([XText], optional): the type of words to count. Defaults to ``None``.

    Raises:
        ValueError: if :paramref:`~.create_range_with_left_right.l_cursor` is missing.
        ValueError: if :paramref:`~.create_range_with_left_right.r_cursor` is missing.

    Returns:
        object: a new instance of a TextCursor which is located at the combined left and right TextRanges.
    """
    if o_text is None:
        o_doc: GenericTextDocument = ooo_util.get_xModel()
        o_text = o_doc.getText()
    cursor = o_text.createTextCursorByRange(l_cursor)
    cursor.gotoRange(r_cursor, expand)
    return cursor


def is_anything_selected(o_doc: Optional[XModel] = None) -> bool:
    """
    Determine if anything is selected.

    Keyword Args:
        o_doc (XModel, optional): current document

    Returns:
        bool: ``True`` if anything in the document is selected: Otherwise, ``False``
    """
    if o_doc is None:
        o_doc = ooo_util.get_xModel()

    o_selections: XIndexAccess = o_doc.getCurrentSelection()
    if not o_selections:
        return False

    count = int(o_selections.getCount())
    if count == 0:
        return False
    elif count > 1:
        return True
    else:
        # There is only one selection so obtain the first selection
        o_sel: XTextRange = o_selections.getByIndex(0)
        o_text: XText = o_sel.getText()
        # Create a text cursor that covers the range and then see if it is collapsed
        o_cursor = o_text.createTextCursorByRange(o_sel)  # XTextCursor
        if not o_cursor.isCollapsed():
            return True

    return False


def doc_select_all(o_doc: Optional[XModel] = None):
    """
    Select all in the current document using Dispatch

    Args:
        o_doc (XModel, optional): current document

    See Also:
        :py:func:`select_view_all`
    """
    if o_doc is None:
        o_doc = ooo_util.get_xModel()
    document: "XDispatchProvider" = o_doc.getCurrentController().getFrame()
    dispatcher: "DispatchHelper" = ooo_util.create_uno_service(
        "com.sun.star.frame.DispatchHelper"
    )
    props: Tuple[PropertyValue] = ()
    dispatcher.executeDispatch(document, ".uno:SelectAll", "", 0, props)


def get_view_cursor(o_doc: Optional[XModel] = None) -> XTextViewCursor:
    """
    Gets current view cursor which is a XTextViewCursor

    Args:
        o_doc (XModel, optional): current document

    Returns:
        object: View Cursor
    """
    if o_doc is None:
        o_doc = ooo_util.get_xModel()
    # https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextViewCursor.html
    frame: XFrame = o_doc.getCurrentController().getFrame()
    current_controler: XTextViewCursorSupplier = frame.getController()
    view_cursor = current_controler.getViewCursor()
    return view_cursor

def select_view_clear(o_doc: Optional[XModel] = None) -> None:
    """
    Clears current view cursor which is a XTextViewCursor.

    Args:
        o_doc (XModel, optional): current document

    See Also:
        :py:func:`doc_select_all`
    """
    if o_doc is None:
        o_doc = ooo_util.get_xModel()
    vc = get_view_cursor(o_doc=o_doc)
    if not vc.isCollapsed():
        vc.collapseToStart()

def select_view_all(o_doc: Optional[XModel] = None) -> None:
    """
    Select all for the current document using view cursor

    Args:
        o_doc (XModel, optional): current document

    Raises:
        Exception: if there is a error selecting

    See Also:
        :py:func:`doc_select_all`
    """
    if o_doc is None:
        o_doc = ooo_util.get_xModel()
    select_view_clear(o_doc=o_doc)
    vc = get_view_cursor(o_doc=o_doc)
    try:
        vc.setVisible(False)
        vc.gotoStart(False)
        vc.gotoEnd(True)
    except Exception as e:
        raise e
    finally:
        if not vc.isVisible():
            vc.setVisible(True)

def create_cursor(o_doc: Optional[XModel] = None, text: Optional[XText] = None) -> XTextCursor:
    """
    Gets a new TextCursor

    Args:
        doc (XModel, optional): current document
        text (XText, optional): Text object to create a TextCursor for.

    Returns:
        object: a new instance of a TextCursor service which can be used to travel in the given text context.

    Note:
        If ``text`` is passed in as an argument then ``doc`` will be ignored.
    """
    if text is None:
        if o_doc is None:
            o_doc: XText = ooo_util.get_xModel()
        text = o_doc.getText()
    return text.createTextCursor()

def get_cursor_start(o_doc: Optional[XModel] = None, text: Optional[XText] = None) -> XTextRange:
    """
    Gets the cursor for the start of the document.

    Args:
        doc (XModel, optional): current documen
        text (XText, optional):Text object to create a TextCursor for.

    Returns:
        object: a new instance of a TextCursor service which can be used to travel in the given text context.

    Note:
        If ``text`` is passed in as an argument then ``doc`` will be ignored.

    See Also:
        :py:func:`get_cursor_end`
    """
    if text is None:
        if o_doc is None:
            o_doc: XText = ooo_util.get_xModel()
        text = o_doc.getText()
    return text.getStart()

def get_cursor_end(o_doc: Optional[XModel] = None, text: Optional[XText] = None) -> XTextRange:
    """
    Gets the cursor for the enf of the document.

    Args:
        doc (XModel, optional): current document
        text (XText, optional): Text object to create a TextCursor for.

    Returns:
        object: a new instance of a TextCursor service which can be used to travel in the given text context.

    Note:
        If ``text`` is passed in as an argument then ``doc`` will be ignored.

    See Also:
        :py:func:`get_cursor_start`
    """
    if text is None:
        if o_doc is None:
            o_doc: XText = ooo_util.get_xModel()
        text = o_doc.getText()
    return text.getEnd()


def get_selected_text(o_doc: Optional[XModel] = None) -> Union[XText, None]:
    """
    Gets the xText for current selection

    Arguments:
        o_doc (XModel, optional): current document

    Returns:
        Union[XText, None]: If no selection is made then None is returned; Otherwise, XText is returned.
    """
    if o_doc is None:
        o_doc = ooo_util.get_xModel()
    o_selections: XIndexAccess = o_doc.getCurrentSelection()
    if not o_selections:
        return None
    count = int(o_selections.getCount())
    if count == 0:
        return None
    o_sel: XText = o_selections.getByIndex(0)
    return o_sel.getText()

def select_next_word() -> None:
    """
    Select the word right from the current curor position.

    Note:
        See capitalisePython LibreOffice Python Example code
    """
    frame = ooo_util.get_xframe()
    x_selection_supplier: XSelectionSupplier = frame.getController()

    # see section 7.5.1 of developers' guide
    x_index_access: XIndexAccess = x_selection_supplier.getSelection()
    x_text_range: XTextRange = x_index_access.getByIndex(0)

    # get the XWordCursor and make a selection!
    xText = x_text_range.getText()
    x_word_cursor: XWordCursor = xText.createTextCursorByRange(x_text_range)

    if not x_word_cursor.isStartOfWord():
        x_word_cursor.gotoStartOfWord(False)

    x_word_cursor.gotoNextWord(True)
    x_selection_supplier.select(x_word_cursor)

# region Private functions
def _get_left_cursor(o_sel: XTextRange, o_text: XText) -> XTextCursor:
    if o_text.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
        o_range = o_sel.getEnd()
    else:
        o_range = o_sel.getStart()
    cursor = o_text.createTextCursorByRange(o_range)
    cursor.goRight(0, False)
    return cursor


def _get_right_cursor(o_sel: XTextRange, o_text: XText) -> XTextCursor:
    if o_text.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
        o_range = o_sel.getStart()
    else:
        o_range = o_sel.getEnd()
    cursor = o_text.createTextCursorByRange(o_range)
    cursor.goLeft(0, False)
    return cursor


# endregion Private functions
