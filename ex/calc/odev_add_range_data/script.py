from range_action import create_array, clear_range
from ooodev.macro.macro_loader import MacroLoader

def fill(*args, **kwargs) -> None:
    with MacroLoader():
        create_array()

def clear(*args, **kwargs) -> None:
    with MacroLoader():
        clear_range()
