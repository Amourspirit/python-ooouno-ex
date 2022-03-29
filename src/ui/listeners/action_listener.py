# coding: utf-8
from typing import TYPE_CHECKING, Callable
import unohelper
from ooo.dyn.awt.x_action_listener import XActionListener

if TYPE_CHECKING:
    from ooo.lo.awt.action_event import ActionEvent


class ActionListener(unohelper.Base, XActionListener):
    """
    Listener class that listens for UNO Actions
    """

    def __init__(self, callback: 'Callable[[ActionEvent], None]'):
        """
        Constructor

        Args:
            callback (callable): Callback function that is called
                when item state change is preformed
        """
        self._callback = callback

    def actionPerformed(self, actionEvent: "ActionEvent") -> None:
        """
        Calls the :paramref:`~.ActionListener.callback` function

        Args:
            actionEvent (object): Action Event
        """
        self._callback()
