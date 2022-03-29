# coding: utf-8
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ooo.lo.awt.uno_control_dialog_element import UnoControlDialogElement


class ControlWrapper:
    """Simple wrapper class that just helps in accessing control properties"""
    def __init__(
        self, name: str, ctl: "UnoControlDialogElement", x: int, y: int
    ) -> None:
        """
        Constructor

        Args:
            name (str): Control Name
            ctl (UnoControlDialogElement): Dialog Element
            x (int): control x coordinate
            y (int): control y coordinate
        """
        self._name = name
        self.ctl = ctl
        self._x = x
        self._y = y

    @property
    def height(self) -> int:
        """specifies the height of the control"""
        return self.ctl.Height

    @height.setter
    def height(self, value: int) -> None:
        self.ctl.Height = value

    @property
    def width(self) -> int:
        """specifies the width of the control"""
        return self.ctl.Width

    @width.setter
    def width(self, value: int) -> None:
        self.ctl.Width = value

    @property
    def name(self) -> int:
        """specifies the name of the control"""
        return self._name

    @name.setter
    def name(self, value: int) -> None:
        self._name = value

    @property
    def x(self) -> int:
        """
        X-coordinate

        :getter: Gets this x-coordinate
        :setter: Sets this x-coordinate
        """
        return self._x

    @x.setter
    def x(self, value: int):
        self._x = value

    @property
    def y(self) -> int:
        """
        Y-coordinate

        :getter: Gets this y-coordinate
        :setter: Sets this y-coordinate
        """
        return self._y

    @y.setter
    def y(self, value: int):
        self._y = value

    @property
    def bottom(self) -> int:
        """
        Gets the bottom positon of this control

        Returns:
            int: Control bottom postion
        """
        return self._y + self.height

    @property
    def right(self) -> int:
        """
        Gets the right postiion of this control

        Returns:
            int: Control right postion
        """
        return self._x + self.width
