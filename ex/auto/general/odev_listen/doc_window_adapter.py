# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING, Any, cast

from ooodev.events.args.event_args import EventArgs
from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.loader import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI

from com.sun.star.awt import XWindow


if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject
# endregion Imports

# region DocWindow Class


class DocWindowAdapter:
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        self.doc = Write.create_doc(loader=loader)

        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to assign the method to class fields and use them to add the event callbacks.
        self._fn_on_disposing = self.on_disposing
        self._fn_on_window_activated = self.on_window_activated
        self._fn_on_window_closed = self.on_window_closed
        self._fn_on_window_closing = self.on_window_closing
        self._fn_on_window_deactivated = self.on_window_deactivated
        self._fn_on_window_minimized = self.on_window_minimized
        self._fn_on_window_normalized = self.on_window_normalized
        self._fn_on_window_opened = self.on_window_opened

        # Assigning TopWindowEvents to class is important,
        # if not assigned then it may get garbage collected after class __init__() is called.
        # By setting add_window_listener=True The TopWindowEvents instance creates a window listener and attaches itself to it.
        self._top_events = TopWindowEvents(add_window_listener=True)
        # The Listener could be accessed via the property: self._top_events.events_listener_top_window
        # Now that _top_events has a listener attached to it we can simply subscribe to it events
        self._top_events.add_event_top_window_events_disposing(self._fn_on_disposing)
        self._top_events.add_event_window_opened(self._fn_on_window_opened)
        self._top_events.add_event_window_activated(self._fn_on_window_activated)
        self._top_events.add_event_window_closed(self._fn_on_window_closed)
        self._top_events.add_event_window_closing(self._fn_on_window_closing)
        self._top_events.add_event_window_deactivated(self._fn_on_window_deactivated)
        self._top_events.add_event_window_minimized(self._fn_on_window_minimized)
        self._top_events.add_event_window_normalized(self._fn_on_window_normalized)
        self._top_events.add_event_window_opened(self._fn_on_window_opened)

        GUI.set_visible(True, self.doc)
        # triggers 2 opened and 2 activated events

    def on_window_opened(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        """is invoked when a window is activated."""
        event = cast("EventObject", event_args.event_data)
        print("WA: Opened")
        x_win = Lo.qi(XWindow, event.Source)
        GUI.print_rect(x_win.getPosSize())

    def on_window_activated(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        """is invoked when a window is activated."""
        print("WA: Activated")
        print(f"  Title bar: {GUI.get_title_bar()}")

    def on_window_deactivated(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        """is invoked when a window is deactivated."""
        print("WA: Minimized")

    def on_window_minimized(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        """is invoked when a window is iconified."""
        print("WA:  De-activated")

    def on_window_normalized(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        """is invoked when a window is deiconified."""
        print("WA: Normalized")

    def on_window_closing(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        """
        is invoked when a window is in the process of being closed.

        The close operation can be overridden at this point.
        """
        print("WA: Closing")

    def on_window_closed(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        """is invoked when a window has been closed."""
        self.closed = True
        print("WA: Closed")

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
        print("WA: Disposing")


# endregion DocWindow Class
