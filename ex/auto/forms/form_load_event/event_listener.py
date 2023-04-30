from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.adapter.adapter_base import AdapterBase, GenericArgs
from com.sun.star.document import DocumentEvent

from com.sun.star.document import XDocumentEventListener
from com.sun.star.document import XEventListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class DocumentEventListener(AdapterBase, XEventListener, XDocumentEventListener):
    """
    Allows notificaton of events happening in an OfficeDocument

    See Also:
        - `API XDocumentEventListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XDocumentEventListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    def documentEventOccured(self, event: DocumentEvent) -> None:
        """
        Is called whenever a document event occurred
        """
        self._trigger_event("documentEventOccured", event)

    def notifyEvent(self, event: EventObject) -> None:
        """
        is called whenever a document event (see EventObject) occurs
        """
        self._trigger_event("notifyEvent", event)

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
