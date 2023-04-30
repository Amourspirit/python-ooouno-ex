from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.adapter.adapter_base import AdapterBase, GenericArgs

from com.sun.star.form import XLoadListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class FormLoadListener(AdapterBase, XLoadListener):
    """
    Receives load-related events from a loadable object.

    The interface is typically implemented by data-bound components, which want to listen to the data source that contains their database form.

    See Also:
        `API XLoadListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1form_1_1XLoadListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    def loaded(self, event: EventObject) -> None:
        """
        is invoked when the object has successfully connected to a datasource.
        """
        self._trigger_event("loaded", event)

    def reloaded(self, event: EventObject) -> None:
        """
        is invoked when the object has been reloaded.
        """
        self._trigger_event("reloaded", event)

    def reloading(self, event: EventObject) -> None:
        """
        is invoked when the object is about to be reloaded.

        Components may use this to stop any other event processing related to the event source until they get the reloaded event.
        """
        self._trigger_event("reloading", event)

    def unloaded(self, event: EventObject) -> None:
        """
        is invoked after the object has disconnected from a datasource.
        """
        self._trigger_event("unloaded", event)

    def unloading(self, event: EventObject) -> None:
        """
        is invoked when the object is about to be unloaded.

        Components may use this to stop any other event processing related to the event source before the object is unloaded.
        """
        self._trigger_event("unloading", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)
