# coding: utf-8
from typing import TYPE_CHECKING, Callable
import unohelper
from ooo.dyn.awt.tab.x_tab_page_container_listener import XTabPageContainerListener

if TYPE_CHECKING:
    from ooo.lo.awt.tab.tab_page_activated_event import TabPageActivatedEvent


class TabPageContainerListener(unohelper.Base, XTabPageContainerListener):
    """
    Listener class that listens for UNO Item State
    """

    def __init__(self, callback: "Callable[[TabPageActivatedEvent], None]"):
        """
        Inits ItemListener

        Args:
            callback (callable): Callback function that is called
                when item state change is preformed
        """
        self._callback = callback

    def tabPageActivated(self, tabPageActivatedEvent: "TabPageActivatedEvent") -> None:
        """
        Calls the :paramref:`~.ItemListener.callback` function

        Args:
            aEvent (object): Item Event
        """
        self._callback(tabPageActivatedEvent)
