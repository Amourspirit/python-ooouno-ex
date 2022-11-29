# region Imports
from __future__ import annotations
import time
import sys
from typing import TYPE_CHECKING, Any, cast

from ooodev.adapter.frame.terminate_listener import TerminateListener, GenericArgs
from ooodev.adapter.lang.event_listener import EventListener
from ooodev.events.args.event_args import EventArgs
from ooodev.events.lo_events import Events
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.office.calc import Calc
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject
# endregion Imports

# region DocMonitor Class


class DocMonitor:
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        self.bridge_disposed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        xdesktop = Lo.XSCRIPTCONTEXT.getDesktop()

        # create a new instance of listener.
        # pass GenericArgs with listener arg of self.
        # this will allow for this instance to be passed to events.
        self._term_listener = TerminateListener(trigger_args=GenericArgs(listener=self))
        self._term_listener.on("notifyTermination", DocMonitor.on_notify_termination)
        self._term_listener.on("queryTermination", DocMonitor.on_query_termination)
        self._term_listener.on("disposing", DocMonitor.on_disposing)

        # using an event is redundant here and is included for example purposes.
        # below a listener is attached to Lo.birdge that does the same job.
        self.events = Events(source=self)
        self.events.on(LoNamedEvent.BRIDGE_DISPOSED, DocMonitor.on_disposed)

        # attach a listener to the bridge connection that gets notified if
        # office bridge connection terminates unexpectly.
        # Lo.bridge is not available if a script is run as a macro.
        self._bridge_listen = EventListener(trigger_args=GenericArgs(listener=self))
        self._bridge_listen.on("disposing", DocMonitor.on_disposing_bridge)
        Lo.bridge.addEventListener(self._bridge_listen)

        self.doc = Calc.create_doc(loader=loader)

        GUI.set_visible(True, self.doc)

    @staticmethod
    def on_notify_termination(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """
        is called when the master environment is finally terminated.

        No veto will be accepted then.
        """
        print("TL: Finished Closing")
        dm = cast(DocMonitor, kwargs["listener"])
        dm.bridge_disposed = True
        dm.closed = True

    @staticmethod
    def on_query_termination(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """
        is called when the master environment (e.g., desktop) is about to terminate.

        Termination can be intercepted by throwing TerminationVetoException.
        Interceptor will be the new owner of desktop and should call XDesktop.terminate() after finishing his own operations.

        Raises:
            TerminationVetoException: ``TerminationVetoException``
        """
        print("TL: Starting Closing")

    @staticmethod
    def on_disposing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """
        gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including XComponent.removeEventListener() ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at XComponent.
        """

        # don't expect Disposing to print if script ends due to closing.
        # script will stop before dispose is called
        print("TL: Disposing")

    @staticmethod
    def on_disposing_bridge(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        # do not try and exit script here.
        # for some reason when office triggers this method calls such as:
        # raise SystemExit(1)
        # does not exit the script
        print("BR: Office bridge has gone!!")

    @staticmethod
    def on_disposed(source: Any, event: EventArgs) -> None:
        # just another way of knowing when bridge is gone.
        print("LO: Office bridge has gone!!")
        dm = cast(DocMonitor, event.event_source)
        dm.bridge_disposed = True

# endregion DocMonitor Class
