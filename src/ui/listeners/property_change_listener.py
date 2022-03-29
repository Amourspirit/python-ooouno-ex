# coding: utf-8
from typing import TYPE_CHECKING, Callable
import unohelper
from ooo.dyn.beans.x_property_change_listener import XPropertyChangeListener
if TYPE_CHECKING:
    from ooo.lo.beans.property_change_event import PropertyChangeEvent


class PropertyChangeListener(unohelper.Base, XPropertyChangeListener):
    """
    Listener class that listens for UNO Item State
    """

    def __init__(self, callback: 'Callable[[PropertyChangeEvent], None]'):
        """
        Inits ItemListener

        Args:
            callback (callable): Callback function that is called
                when item state change is preformed
        """
        self._callback = callback
    
    def propertyChange(self, evt: 'PropertyChangeEvent') -> None:
        self._callback(evt)
