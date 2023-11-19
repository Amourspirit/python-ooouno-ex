# region Imports
from __future__ import annotations
from typing import Any

from ooodev.adapter.frame.terminate_events import TerminateEvents
from ooodev.adapter.lang.event_events import EventEvents
from ooodev.events.args.event_args import EventArgs
from ooodev.events.lo_events import Events
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.office.calc import Calc
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

# endregion Imports

# region DocMonitor Class


class DocMonitor:
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        self.bridge_disposed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        _ = Lo.XSCRIPTCONTEXT.getDesktop()

        self._set_internal_events()

        self.doc = Calc.create_doc(loader=loader)

        GUI.set_visible(True, self.doc)

    def _set_internal_events(self):
        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to assign the method to class fields and use them to add the event callbacks.
        self._fn_on_notify_termination = self.on_notify_termination
        self._fn_on_query_termination = self.on_query_termination
        self._fn_on_disposing = self.on_disposing
        self._fn_on_disposing_bridge = self.on_disposing_bridge
        self._fn_on_disposed = self.on_disposed

        # create a new instance of listener events.
        # By default TerminateEvents will create a listener and attach itself to the current XDesktop instance.
        self._term_events = TerminateEvents()
        # The Listener could be accessed via the property: self._term_events.events_listener_terminate
        self._term_events.add_event_notify_termination(self._fn_on_notify_termination)
        self._term_events.add_event_query_termination(self._fn_on_query_termination)
        self._term_events.add_event_terminate_events_disposing(self._fn_on_disposing)

        # using an event is redundant here and is included for example purposes.
        # below a listener is attached to Lo.bridge that does the same job.
        self.events = Events(source=self)
        self.events.on(LoNamedEvent.BRIDGE_DISPOSED, self._fn_on_disposed)

        # attach a listener to the bridge connection that gets notified if
        # office bridge connection terminates unexpected.
        # Lo.bridge is not available if a script is run as a macro.
        self._bridge_events = EventEvents()
        self._bridge_events.add_event_disposing(self._fn_on_disposing)
        Lo.bridge.addEventListener(self._bridge_events.events_listener_event)

    def on_notify_termination(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """
        is called when the master environment is finally terminated.

        No veto will be accepted then.
        """
        print("TL: Finished Closing")
        self.bridge_disposed = True
        self.closed = True

    def on_query_termination(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """
        is called when the master environment (e.g., desktop) is about to terminate.

        Termination can be intercepted by throwing TerminationVetoException.
        Interceptor will be the new owner of desktop and should call XDesktop.terminate() after finishing his own operations.

        Raises:
            TerminationVetoException: ``TerminationVetoException``
        """
        print("TL: Starting Closing")

    def on_disposing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
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

    def on_disposing_bridge(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        # do not try and exit script here.
        # for some reason when office triggers this method calls such as:
        # raise SystemExit(1)
        # does not exit the script
        print("BR: Office bridge has gone!!")

    def on_disposed(self, source: Any, event_args: EventArgs) -> None:
        # just another way of knowing when bridge is gone.
        print("LO: Office bridge has gone!!")
        self.bridge_disposed = True


# endregion DocMonitor Class
