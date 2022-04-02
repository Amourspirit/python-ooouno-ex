# coding: utf-8
from typing import Union
from .interface import IControllerMultiSyntax, IViewMultiSyntax, IModelMuiltiSyntax, IControllerVoid, IControllerSyntax
from .enums import TabEnum
from .controller_syntax import ControllerSyntax
from .controller_void import ControllerVoid


class MultiSyntaxController(IControllerMultiSyntax):
    def __init__(self, model: IModelMuiltiSyntax, view: IViewMultiSyntax):
        self._model = model
        self._view = view
        self._title = "Syntax"
        self._selected_text = 'Hello world'
        self._selected_tab = TabEnum.SYNTAX
        self._controller_syntax = ControllerSyntax(self, self._view)
        self._controller_void = ControllerVoid(self, self._view)

    def start(self):
        self._view.setup(controller=self)
        self._view.active_tab = self._selected_tab
        self._view.show()


    # region Handler Methods

    def handle_click_cancel(self):
        """Cancel Event handleer"""
        self._view.exit()

    def handle_click_ok(self):
        try:
            if self.selected_tab == TabEnum.SYNTAX:
                self._controller_syntax.write()
            elif self.selected_tab == TabEnum.VOID:
                self._controller_void.write()
            else:
                raise NotImplementedError(
                    "MultiSyntaxController.handle_click_cancel()")
        except Exception as e:
            print(e)
            raise e
        finally:
            self._view.exit()
    # endregion Handler Methods
    # region Properties

    @property
    def title(self) -> str:
        """Gets/sets title"""
        return self._title

    @title.setter
    def title(self, value: str):
        self._title = value

    @property
    def selected_text(self) -> Union[str, None]:
        """Gets word data"""
        return self._selected_text

    @selected_text.setter
    def selected_text(self, value: str):
        self._selected_text = value
        self._view.refresh()


    @property
    def model(self) -> IControllerMultiSyntax:
        """
        Specifies model_syntax value
        """
        return self._model

    @property
    def view(self) -> IViewMultiSyntax:
        """Gets view value"""
        return self._view

    @property
    def selected_tab(self) -> TabEnum:
        """Specifies selected_tab

            :getter: Gets selected_tab value.
            :setter: Sets selected_tab value.
        """
        return self._selected_tab

    @selected_tab.setter
    def selected_tab(self, value: TabEnum):
        self._selected_tab = value
        self._view.refresh()

    @property
    def controller_syntax(self) -> IControllerSyntax:
        """Gets controller_syntax value"""
        return self._controller_syntax

    @property
    def controller_void(self) -> IControllerVoid:
        """Gets controller_void value"""
        return self._controller_void
    # endregion Properties
