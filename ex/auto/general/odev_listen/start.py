#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import time
import sys
from typing import TYPE_CHECKING

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.office.write import Write
from ooodev.listeners.x_top_window_adapter import XTopWindowAdapter

# see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/listeners/x_top_window_adapter.html
from com.sun.star.awt import XExtendedToolkit
from com.sun.star.awt import XWindow


if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject
# endregion Imports

# region DocWindow Class


class DocWindow(XTopWindowAdapter):
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        # assigning tk to class is important.
        # if not assigned then tk goes out of scope after class __init__() is callde
        # and dispose is called right after __init__()
        self.tk = Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")
        if self.tk is not None:
            self.tk.addTopWindowListener(self)

        self.doc = Write.create_doc(loader=loader)

        GUI.set_visible(True, self.doc)
        # triggers 2 opened and 2 activated events

    def windowOpened(self, event: EventObject) -> None:
        """is invoked when a window is activated."""
        print("WL: Opened")
        xwin = Lo.qi(XWindow, event.Source)
        GUI.print_rect(xwin.getPosSize())

    def windowActivated(self, event: EventObject) -> None:
        """is invoked when a window is activated."""
        print("WL: Activated")
        print(f"  Titile bar: {GUI.get_title_bar()}")

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

# region main


def main_loop() -> None:
    # https://stackoverflow.com/a/8685815/1171746
    dw = DocWindow()

    # delay in seconds
    delay = 1.5

    # start run min and max to raise listen events
    time.sleep(delay) # wait delay amount of seconds
    for _ in range(3):
        time.sleep(delay)
        GUI.minimize(dw.doc)
        time.sleep(delay)
        GUI.maximize(dw.doc)
    
    # check an see if user passed in a auto terminate option
    if len(sys.argv) > 1:
        if str(sys.argv[1]).casefold() in ('t', 'true', 'y', 'yes'):
            Lo.delay(delay)
            Lo.close_office()

    # stop run min and max to raise listen events

    # while Writer is open, keep running the script unless specifically ended by user
    while 1:
        if dw.closed is True: # wait for windowClosed event to be raised
            print("\nExiting by document close.\n")
            break
        time.sleep(0.1)


if __name__ == "__main__":
    print("Press 'ctl+c' to exit script early.")
    try:
        main_loop()
    except KeyboardInterrupt:
        # ctrl+c exitst the script earily
        print("\nExiting by user request.\n", file=sys.stderr)
        sys.exit(0)

# endregion main
