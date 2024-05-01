from __future__ import annotations
from typing import Tuple
from urllib.parse import parse_qs
import contextlib
import uno
import unohelper
from com.sun.star.frame import XDispatchProviderInterceptor
from com.sun.star.frame import XDispatchProvider
from com.sun.star.frame import XDispatch
from com.sun.star.util import URL
from com.sun.star.frame import DispatchDescriptor
from dispatch_convert_cell_url import DispatchConvertCellUrl


class DispatchProviderInterceptor(unohelper.Base, XDispatchProviderInterceptor):

    def __init__(self):
        self._master = None
        self._slave = None

    # def _convert_query_to_dict(self, query: str):
    #     return parse_qs(query)

    def _convert_query_to_dict(self, query: str):
        query_dict = parse_qs(query)
        return {k: v[0] for k, v in query_dict.items()}

    def getMasterDispatchProvider(self) -> XDispatchProvider:
        """
        access to the master XDispatchProvider of this interceptor
        """
        return self._master

    def getSlaveDispatchProvider(self) -> XDispatchProvider:
        """
        access to the slave XDispatchProvider of this interceptor
        """
        return self._slave

    def setMasterDispatchProvider(self, new_supplier: XDispatchProvider) -> None:
        """
        sets the master XDispatchProvider, which may forward calls to its XDispatchProvider.queryDispatch() to this dispatch provider.
        """
        self._master = new_supplier

    def setSlaveDispatchProvider(
        self, new_dispatch_provider: XDispatchProvider
    ) -> None:
        """
        sets the slave XDispatchProvider to which calls to XDispatchProvider.queryDispatch() can be forwarded under control of this dispatch provider.
        """
        self._slave = new_dispatch_provider

    def queryDispatch(
        self, url: URL, target_frame_name: str, search_flags: int
    ) -> XDispatch | None:
        """
        Searches for an XDispatch for the specified URL within the specified target frame.
        """
        if url.Protocol == "slot:":
            # not really sure if this is necessary.
            return None

        if url.Main == ".uno:ooodev.calc.menu.convert.url":
            # print(
            #     f"Complete = {url.Complete}",
            #     f"Protocol = {url.Protocol}",
            #     f"Main = {url.Main}",
            #     f"Mark = {url.Mark}",
            #     f"Name = {url.Name}",
            #     f"Path = {url.Path}",
            #     f"User = {url.User}",
            #     f"Server = {url.Server}",
            #     f"Port = {url.Port}",
            #     f"Arguments = {url.Arguments}",
            #     sep="; ",
            # )
            # print(url)
            with contextlib.suppress(Exception):
                args = self._convert_query_to_dict(url.Arguments)
                return DispatchConvertCellUrl(sheet=args["sheet"], cell=args["cell"])
        return self._slave.queryDispatch(url, target_frame_name, search_flags)

    def queryDispatches(
        self, requests: Tuple[DispatchDescriptor, ...]
    ) -> Tuple[XDispatch, ...]:
        """
        Actually this method is redundant to XDispatchProvider.queryDispatch() to avoid multiple remote calls.

        It's not allowed to pack it - because every request must match to its real result. Means: don't delete NULL entries inside this list.
        """
        pass
