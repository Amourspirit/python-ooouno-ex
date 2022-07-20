#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import time
import sys
import types
from typing import TYPE_CHECKING

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.listeners.x_terminate_adapter import XTerminateAdapter


if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject
# endregion Imports

# region DocMonitor Class


class DocMonitor:
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        xdesktop = Lo.XSCRIPTCONTEXT.getDesktop()
        
        # create a new instance of adapter. Note that adapter methods all pass.
        term_adapter = XTerminateAdapter()
        
        # reassign the method we want to use from XTerminateAdapter instance in a pythonic way.
        term_adapter.notifyTermination = types.MethodType(self.notify_termination, term_adapter)
        term_adapter.queryTermination = types.MethodType(self.query_termination, term_adapter)
        term_adapter.disposing = types.MethodType(self.disposing, term_adapter)
        xdesktop.addTerminateListener(term_adapter)
        
        self.doc = Calc.create_doc(loader=loader)

        GUI.set_visible(True, self.doc)


    def notify_termination(self, src: XTerminateAdapter, event: EventObject) -> None:
        """
        is called when the master environment is finally terminated.

        No veto will be accepted then.
        """
        print("TL: Finished Closing")
        self.closed = True
    
    def query_termination(self, src: XTerminateAdapter, event: EventObject) -> None:
        """
        is called when the master environment (e.g., desktop) is about to terminate.

        Termination can be intercepted by throwing TerminationVetoException.
        Interceptor will be the new owner of desktop and should call XDesktop.terminate() after finishing his own operations.

        Raises:
            TerminationVetoException: ``TerminationVetoException``
        """
        print("TL: Starting Closing")
        
        
    def disposing(self, src: XTerminateAdapter, event: EventObject) -> None:
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

# endregion DocMonitor Class

# region main


def main_loop() -> None:
    # https://stackoverflow.com/a/8685815/1171746
    dw = DocMonitor()

    # check an see if user passed in a auto terminate option
    if len(sys.argv) > 1:
        if str(sys.argv[1]).casefold() in ('t', 'true', 'y', 'yes'):
            Lo.delay(5000)
            Lo.close_office()

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
