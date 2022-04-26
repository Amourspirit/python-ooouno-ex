from __future__ import annotations
from pprint import pprint
from typing import Union, TYPE_CHECKING
import scriptforge as SF
from . import sudoku
import threading

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheet
    from com.sun.star.util import XProtectable

    _SHEET = Union[XSpreadsheet, XProtectable]

_game_board = None
_solution = None
_STYLE_EMPTY_CELL = "Accent 2"
_STYLE_FIXED_CELL = "Accent 1"
_STYLE_GOOD_CELL = "Good"
_STYLE_BAD_CELL = "Bad"
_STYLE_HINT_CELL = "Neutral"
_SHEET_NAME = "sudoku"


def run_in_thread(fn):
    """
    Decorator function.
    
    Run any function in thread

    Args:
        fn (Any): Python function (macro)
    """

    def run(*k, **kw):
        t = threading.Thread(target=fn, args=k, kwargs=kw)
        t.start()
        return t

    return run


def _is_board() -> bool:
    global _game_board
    return _game_board is not None


def _create_matrix_single_solve() -> None:
    global _game_board, _solution
    _game_board = []
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    board = sudoku.generate_single_solve_board()
    _solution = sudoku.get_current_solution()
    # values = []
    for row in board:
        _game_board.append([value or None for value in row])

    doc.SetArray("A1", _game_board)


def reset_board() -> None:
    global _game_board
    if not _is_board():
        return
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    sht = doc.XSpreadsheet(_SHEET_NAME)
    _unprotect_sheet(sht)
    doc.SetArray("A1", _game_board)
    _set_board_styles()
    _protect_sheet(sht)


def _unprotect_sheet(sht: _SHEET) -> None:
    if sht.isProtected():
        sht.unprotect("")


def _protect_sheet(sht: _SHEET) -> None:
    if not sht.isProtected():
        sht.protect("")


def _is_origin(x: int, y: int) -> bool:
    # zero-based index
    # get if cell is part of original gameboard.
    # print("x", x, "y", y)
    global _game_board
    result = False
    try:
        i = _game_board[x][y]
        result = isinstance(i, int)
    except Exception as e:
        # print(e)
        pass
    return result


# @run_in_thread
def _set_board_styles() -> None:
    # sets board styles after generation of a new board.
    global _game_board
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.SetCellStyle("A1:I9", _STYLE_EMPTY_CELL)
    for i, row in enumerate(_game_board):
        for j, val in enumerate(row):
            if val is not None:
                col_name = doc.GetColumnName(j + 1)
                doc.SetCellStyle(
                    f"{col_name}{i + 1}:{col_name}{i + 1}", _STYLE_FIXED_CELL
                )

def _style_empty_cell(doc: SF.SFDocuments.SF_Calc, cell: str) -> None:
    doc.SetCellStyle(cell, _STYLE_EMPTY_CELL)


def _style_good_cell(doc: SF.SFDocuments.SF_Calc, cell: str) -> None:
    doc.SetCellStyle(cell, _STYLE_GOOD_CELL)


def _style_bad_cell(doc: SF.SFDocuments.SF_Calc, cell: str) -> None:
    doc.SetCellStyle(cell, _STYLE_BAD_CELL)


def _style_hint_cell(doc: SF.SFDocuments.SF_Calc, cell: str) -> None:
    doc.SetCellStyle(cell, _STYLE_HINT_CELL)

def _possible(row: int, col: int, num: int) -> bool:
    # working with a single solution board.
    # makes checking for possible in this case.
    global _solution
    # print(f"checking for possible: Row:{row}, Column:{col} for number: {num}")
    return _solution[row][col] == num


def set_number(num: int) -> None:
    # set a cell to a number. num range 1-9
    if not _is_board():
        return
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    cur_selection: str = doc.CurrentSelection
    # Gets address of first cell in the selection
    cell = doc.Offset(cur_selection, 0, 0, 1, 1)
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")

    if doc.LastRow(cell) > 9 or doc.LastColumn(cell) > 9:
        bas.MsgBox(f"Please select a game board cell.", title="Outside range")
        return
    x = doc.FirstRow(cell) - 1
    y = doc.FirstColumn(cell) - 1
    # see if x,y is an origin from board generation.
    if _is_origin(x, y):
        bas.MsgBox("This value is locked", title="Original value")
        return
    sht = doc.XSpreadsheet(_SHEET_NAME)
    _unprotect_sheet(sht)
    doc.SetValue(cell, num)
    if _possible(x, y, num):
        _style_good_cell(doc, cell)
    else:
        _style_bad_cell(doc, cell)
    _protect_sheet(sht)

def hint() -> None:
    # gets value from solution for the current cell and displays it.
    if not _is_board():
        return
    global _solution
    if _solution is None:
        return
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    cur_selection: str = doc.CurrentSelection
    # Gets address of first cell in the selection
    cell = doc.Offset(cur_selection, 0, 0, 1, 1)
    if doc.LastRow(cell) > 9 or doc.LastColumn(cell) > 9:
        return
    x = doc.FirstRow(cell) - 1
    y = doc.FirstColumn(cell) - 1
    # see if x,y is an origin from board generation.
    if _is_origin(x, y):
        return
    # print("x", x, "y", y)
    # pprint(_solution)
    hint_val = _solution[x][y]
    # print("Hint", hint_val)

    sht = doc.XSpreadsheet(_SHEET_NAME)
    _unprotect_sheet(sht)
    doc.SetValue(cell, hint_val)
    _style_hint_cell(doc, cell)
    _protect_sheet(sht)

def clear_cell() -> None:
    # clears current selected cell on board.
    if not _is_board():
        return
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    cur_selection: str = doc.CurrentSelection
    # Gets address of first cell in the selection
    cell = doc.Offset(cur_selection, 0, 0, 1, 1)
    if doc.LastRow(cell) > 9 or doc.LastColumn(cell) > 9:
        return
    x = doc.FirstRow(cell) - 1
    y = doc.FirstColumn(cell) - 1
    # see if x,y is an origin from board generation.
    if _is_origin(x, y):
        return
    sht = doc.XSpreadsheet(_SHEET_NAME)
    _unprotect_sheet(sht)
    doc.SetValue(cell, "")
    _style_empty_cell(doc=doc, cell=cell)
    _protect_sheet(sht)


def new_game() -> None:
    # create a new game.
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    sht = doc.XSpreadsheet(_SHEET_NAME)
    _unprotect_sheet(sht)
    _create_matrix_single_solve()
    _set_board_styles()
    _protect_sheet(sht)
