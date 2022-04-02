# coding: utf-8
from abc import ABC, abstractmethod, abstractproperty
from typing import List, Union
from .enums import TabEnum, TimeEnum, TextEnum, SyntaxEnum, VoidEnum


class IControllerVoid(ABC):
    """Controller Interface Class of MVC for void"""

    @abstractmethod
    def __init__(self, controller: "IControllerMultiSyntax", view: "IViewMultiSyntax"):
        """Constructor"""

    @abstractmethod
    def start(self):
        """Start Controller"""

    @abstractmethod
    def get_list_data(self) -> List[str]:
        """ "Get Void List Data"""

    @abstractmethod
    def write(self):
        """write method"""

    @abstractproperty
    def void(self) -> Union[VoidEnum, None]:
        """Gets/sets syntax enum or None"""

    # endregion properties


class IControllerSyntax(ABC):
    """Controller Interface Class of MVC for void"""

    @abstractmethod
    def __init__(self, controller: "IControllerMultiSyntax", view: "IViewMultiSyntax"):
        """Constructor"""

    @abstractmethod
    def start(self):
        """Start Controller"""

    @abstractmethod
    def get_list_data(self) -> List[str]:
        """ "Get Void List Data"""

    @abstractmethod
    def write(self):
        """write method"""

    @abstractmethod
    def is_syntax(self) -> bool:
        """Gets if Property ``syntax`` is none or a value"""

    # region properties
    @abstractproperty
    def dpv(self) -> bool:
        """Gets/Sets dpv state"""

    @abstractproperty
    def time(self) -> TimeEnum:
        """Gets/Sets dpv time"""

    @abstractproperty
    def text(self) -> TextEnum:
        """Gets/sets prefix suffix state"""

    @abstractproperty
    def syntax(self) -> Union[SyntaxEnum, None]:
        """Gets/sets syntax enum or None"""

    # endregion properties


class IViewMultiSyntax(ABC):
    """View Interface class of MVC for Void"""

    @abstractmethod
    def setup(self, controller: "IControllerMultiSyntax"):
        """Setup main view"""

    @abstractmethod
    def show(self):
        """Show Dialog"""

    @abstractmethod
    def refresh(self):
        """refresh dialog"""

    @abstractmethod
    def exit(self):
        """Exit View"""

    @abstractproperty
    def active_tab(self) -> TabEnum:
        """Gets/sets activte tab"""


class IModelMuiltiSyntax:
    """Model Interface class of MVC for MultiSyntax"""

    @abstractmethod
    def get_syntax_list_data(self) -> List[str]:
        """Get syntax List Data"""

    @abstractmethod
    def get_void_list_data(self) -> List[str]:
        """Get void List Data"""


class IControllerMultiSyntax(ABC):
    """Controller Interface Class of MVC for void"""

    @abstractmethod
    def __init__(self, model: IModelMuiltiSyntax, view: IViewMultiSyntax):
        """Constructor"""

    @abstractmethod
    def start(self):
        """Start Controller"""

    @abstractmethod
    def handle_click_cancel(self):
        """Cancel Event handleer"""

    @abstractmethod
    def handle_click_ok(self):
        """Ok Before Event handleer"""

    # region properties
    @abstractproperty
    def title(self) -> str:
        """Gets/sets title"""

    @abstractproperty
    def selected_text(self) -> str:
        """Gets selected text"""

    @abstractproperty
    def model(self) -> IModelMuiltiSyntax:
        """
        Gets model value
        """

    @abstractproperty
    def view(self) -> IViewMultiSyntax:
        """Gets view value"""

    @abstractproperty
    def selected_tab(self) -> TabEnum:
        """Specifies selected_tab

        :getter: Gets selected_tab value.
        :setter: Sets selected_tab value.
        """

    @abstractproperty
    def controller_syntax(self) -> IControllerSyntax:
        """Specifies Controller for syntax"""

    @abstractproperty
    def controller_void(self) -> IControllerVoid:
        """Specifies Controller for void"""
