from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.adapter.adapter_base import AdapterBase, GenericArgs

from com.sun.star.frame import XLoadEventListener
from com.sun.star.frame import XFrameLoader

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class FrameLoadListener(AdapterBase, XLoadEventListener):
    """
    is used to receive callbacks from an asynchronous frame loader.

    See Also:
        `API XLoadEventListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XLoadEventListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    def loadCancelled(self, loader: XFrameLoader) -> None:
        """
        is called when a frame load is canceled or failed.
        """
        self._trigger_event("loadCancelled", loader)

    def loadFinished(self, loader: XFrameLoader) -> None:
        """
        is called when a new component is loaded into a frame successfully.
        """
        self._trigger_event("loadFinished", loader)

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
