from typing import Any
import counter


def start_dialog(*args, **kwargs) -> None:
    counter.start_dialog()


def btnIncrement(event: Any) -> None:
    counter.updateLabel(event, 1)


def btnDecrement(event: Any) -> None:
    counter.updateLabel(event, -1)
