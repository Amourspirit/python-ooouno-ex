# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING, Any, cast

from ooodev.adapter.awt.top_window_listener import TopWindowListener, EventArgs, GenericArgs
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

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
        # The solution is to create a function that calls the class method and pass that function to the event system.
        # Also the function must be a member of the class so that it is not garbage collected.

        def _on_window_opened(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_opened(source=source, event_args=event_args, *args, **kwargs)

        def _on_window_activated(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_activated(source=source, event_args=event_args, *args, **kwargs)

        def _on_window_deactivated(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_deactivated(source=source, event_args=event_args, *args, **kwargs)

        def _on_window_minimized(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_minimized(source=source, event_args=event_args, *args, **kwargs)

        def _on_window_normalized(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_normalized(source=source, event_args=event_args, *args, **kwargs)

        def _on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_closing(source=source, event_args=event_args, *args, **kwargs)

        def _on_window_closed(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_closed(source=source, event_args=event_args, *args, **kwargs)

        def _on_disposing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_disposing(source=source, event_args=event_args, *args, **kwargs)

        self._fn_on_window_openeed = _on_window_opened
        self._fn_on_window_activated = _on_window_activated
        self._fn_on_window_deactivated = _on_window_deactivated
        self._fn_on_window_minimized = _on_window_minimized
        self._fn_on_window_normalized = _on_window_normalized
        self._fn_on_window_closing = _on_window_closing
        self._fn_on_window_closed = _on_window_closed
        self._fn_on_disposing = _on_disposing

        # assigning TopWindowListener to class is important.
        # if not assigned then tk goes out of scope after class __init__() is called
        # and dispose is called right after __init__()
        self._twl = TopWindowListener()
        self._twl.on("windowOpened", _on_window_opened)
        self._twl.on("windowActivated", _on_window_activated)
        self._twl.on("windowDeactivated", _on_window_deactivated)
        self._twl.on("windowMinimized", _on_window_minimized)
        self._twl.on("windowNormalized", _on_window_normalized)
        self._twl.on("windowClosing", _on_window_closing)
        self._twl.on("windowClosed", _on_window_closed)
        self._twl.on("disposing", _on_disposing)

        GUI.set_visible(True, self.doc)
        # triggers 2 opened and 2 activated events

    def on_window_opened(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """is invoked when a window is activated."""
        event = cast("EventObject", event_args.event_data)
        print("WA: Opened")
        xwin = Lo.qi(XWindow, event.Source)
        GUI.print_rect(xwin.getPosSize())

    def on_window_activated(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """is invoked when a window is activated."""
        print("WA: Activated")
        print(f"  Titile bar: {GUI.get_title_bar()}")

    def on_window_deactivated(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """is invoked when a window is deactivated."""
        print("WA: Minimized")

    def on_window_minimized(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """is invoked when a window is iconified."""
        print("WA:  De-activated")

    def on_window_normalized(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """is invoked when a window is deiconified."""
        print("WA: Normalized")

    def on_window_closing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        """
        is invoked when a window is in the process of being closed.

        The close operation can be overridden at this point.
        """
        print("WA: Closing")

    def on_window_closed(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
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
