import scriptforge as SF
from . import sudoku

_game_board = None

def create_matrix_single_solve() -> None:
    global _game_board
    _game_board = []
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    board = sudoku.generate_single_solve_board()
    # values = []
    for row in board:
        _game_board.append([value or None for value in row])
    
    doc.SetArray("A1", _game_board)

def set_number(num: int) -> None:
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
    if is_origin(x, y):
        bas.MsgBox("This value is locked", title="Original value")
        return
    doc.SetValue(cell, num)

def is_origin(x: int, y: int) -> bool:
    print("x", x, "y", y)
    global _game_board
    result = False
    try:
        i = _game_board[x][y]
        result = isinstance(i, int)
    except Exception as e:
        print(e)
    return result

def set_board_styles() -> None:
    global _game_board
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.SetCellStyle("A1:I9", "Default")
    for i, row in enumerate(_game_board):
        for j, val in enumerate(row):
            if val is not None:
                col_name = doc.GetColumnName(j + 1)
                doc.SetCellStyle(f"{col_name}{i + 1}:{col_name}{i + 1}", "Accent 1")
    
def new_game() -> None:
    create_matrix_single_solve()
    set_board_styles()