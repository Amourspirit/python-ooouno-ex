# coding: utf-8
from typing import TYPE_CHECKING, Callable
import unohelper
from ooo.dyn.awt.x_item_listener import XItemListener

if TYPE_CHECKING:
    from ooo.lo.awt.item_event import ItemEvent


class ItemListener(unohelper.Base, XItemListener):
    """
    Listener class that listens for UNO Item State
    """

    def __init__(self, callback: 'Callable[[ItemEvent], None]'):
        """
        Inits ItemListener

        Args:
            callback (callable): Callback function that is called
                when item state change is preformed
        """
        self._callback = callback

    def itemStateChanged(self, rEvent: 'ItemEvent'):
        """
        Calls the :paramref:`~.ItemListener.callback` function

        Args:
            aEvent (object): Item Event
        """
        self._callback(rEvent)
