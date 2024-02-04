# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING
import unohelper

from ooodev.loader import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI

from com.sun.star.awt import XExtendedToolkit
from com.sun.star.awt import XTopWindowListener
from com.sun.star.awt import XWindow


if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject
# endregion Imports

# region DocWindow Class


class DocWindow(unohelper.Base, XTopWindowListener):
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        self.doc = Write.create_doc(loader=loader)

        # assigning tk to class is important.
        # if not assigned then tk goes out of scope after class __init__() is called
        # and dispose is called right after __init__()
        self.tk = Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")
        if self.tk is not None:
            self.tk.addTopWindowListener(self)

        GUI.set_visible(True, self.doc)
        # triggers 2 opened and 2 activated events

    def windowOpened(self, event: EventObject) -> None:
        """is invoked when a window is activated."""
        print("WL: Opened")
        x_win = Lo.qi(XWindow, event.Source)
        GUI.print_rect(x_win.getPosSize())

    def windowActivated(self, event: EventObject) -> None:
        """is invoked when a window is activated."""
        print("WL: Activated")
        print(f"  Title bar: {GUI.get_title_bar()}")

    def windowDeactivated(self, event: EventObject) -> None:
        """is invoked when a window is deactivated."""
        print("WL: Minimized")

    def windowMinimized(self, event: EventObject) -> None:
        """is invoked when a window is iconified."""
        print("WL:  De-activated")

    def windowNormalized(self, event: EventObject) -> None:
        """is invoked when a window is deiconified."""
        print("WL: Normalized")

    def windowClosing(self, event: EventObject) -> None:
        """
        is invoked when a window is in the process of being closed.

        The close operation can be overridden at this point.
        """
        print("WL: Closing")

    def windowClosed(self, event: EventObject) -> None:
        """is invoked when a window has been closed."""
        if not self.closed:
            print("WL: Closed")
            self.closed = True

    def disposing(self, event: EventObject) -> None:
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
        print("WL: Disposing")


# endregion DocWindow Class
