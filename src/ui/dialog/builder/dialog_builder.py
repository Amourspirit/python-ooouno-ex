# coding: utf-8
from ooodev.utils.lo import Lo
from ooo.dyn.awt.pos_size import PosSize
from typing import cast, Set, Union, TYPE_CHECKING, TypeVar
import uno
from .control_wrapper import ControlWrapper
if TYPE_CHECKING:
    from ooo.lo.awt.uno_control_dialog import UnoControlDialog
    from ooo.lo.awt.uno_control_dialog_model import UnoControlDialogModel
    # from ooo.lo.awt.uno_control_model import UnoControlModel
    from ooo.lo.uno.x_component_context import XComponentContext
    from ooo.lo.uno.x_interface import XInterface
    from ooo.lo.frame.x_frame import XFrame
    from ooo.lo.awt.x_window import XWindow
    from ooo.lo.awt.uno_control_dialog_element import UnoControlDialogElement

T = TypeVar("T")


class DialogBuilder:
    """
    Class that handler creating of a new LibreOffice Dialog Window
    """

    def __init__(self, title: str, width: int, height: int):
        """
        Inits Class instance

        Args:       
            title (str): Title of this dialog
            width (int): Width of this dialog
            height (int): Height of this dialog
        """
        self._title: str = title
        self._width: int = width
        self._height: int = height
        self._x: int = 0
        self._y: int = 0
        self._ctl_names: Set[str] = set()
        # self._ctx = cast('XComponentContext', uno.getComponentContext())
        self._ctx  = Lo.XSCRIPTCONTEXT.getComponentContext()
        self._dialog: 'UnoControlDialog' = self._create(
            "com.sun.star.awt.UnoControlDialog")
        self._dialog_model: 'UnoControlDialogModel' = self._create(
            "com.sun.star.awt.UnoControlDialogModel")  # XControlModel

        # setModel from  com::sun::star::awt::XControlModel
        self._dialog.setModel(self._dialog_model)
        self._dialog.setVisible(False)
        self._dialog.setTitle(title)
        self._dialog.setPosSize(
            self._x, self._y, self._width, self._height, PosSize.SIZE)

    def set_position_size(self, x: int, y: int):
        """
        Sets the position if this dialog

        Args:
            x (int): x coordinate
            y (int): y coordinate
        """
        self._x = x
        self._y = y
        self._dialog.setPosSize(self._x, self._y, 0, 0, PosSize.POS)

    def get_control(self, name: str) -> 'UnoControlDialogElement':
        """
        Get a control of the dialog

        Args:
            name (str): the name of the control to get

        Raises:
            Exception: if control does not exsit.

        Returns:
            object: [description]
        """
        if not name in self._ctl_names:
            raise Exception(
                "Control by the name of '{0}' does not exist".format(name))
        return self._dialog.getControl(name)

    def get_frame(self) -> 'Union[XFrame, None]':
        """
        Gets the current frame if it exist

        Returns:
            Union[XFrame, None]: Current UNO frame if exist; Otherwise, ``None``
        
        See Also:
            `LibreOffice API - XDesktop Interface Reference <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XDesktop.html#a579f3254892c3a0750e92fe0de4e586c>`_
        """
        # dt = theDesktop()
        # frame = dt.getCurrentFrame()
        frame = self._create("com.sun.star.frame.Desktop").getCurrentFrame()
        return frame

    def get_window(self) -> 'Union[XWindow, None]':
        """
        Gets the current window if exist

        Returns:
            Union[XWindow, None]: the container window of this frame if exist; Otherwise, ``None``
        
        See Also:
            `LibreOffice API - XFrame Interface Reference <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XFrame.html#a392191f1033b31f871b6328b1f634530>`_
        """
        frame = self.get_frame()
        window = frame.getContainerWindow() if frame else None
        return window

    def init(self):
        """
        Renders the conrols for this dialog
        """
        window = self.get_window()
        self._dialog.createPeer(self._create(
            "com.sun.star.awt.Toolkit"), window)

    def endExecute(self):
        """
        End this dialog
        """
        self._controls.clear()
        self._dialog.endExecute()

    def dispose(self):
        ''' 
        Dispose of this dialog
        '''
        self._controls.clear()
        self._dialog.dispose()

    def create_control(self, ctl: T) -> T:
        """
        Creates a new Uno Control

        Args:
            name (UnoControlDialogElement): Control that implements UnoControlDialogElement.

        Returns:
            UnoControlDialogElement: Control instance
        """
        return self._create_control(ctl)

    def _create_control(self, ctl: 'Union[str, UnoControlDialogElement]') -> 'UnoControlDialogElement':
        """
        creates a new Uno Control

        Args:
            name (str): uno control name such as ``com.sun.star.awt.UnoControlButtonModel``

        Returns:
            UnoControlDialogElement: Control instance
        """
        if isinstance(ctl, str):
            return self._dialog_model.createInstance(ctl)
        if not hasattr(ctl, '__ooo_full_ns__'):
            raise TypeError(
                f"{self.__class__.__name__}.create_control() `ctl` is not a valid `UnoControlDialogElement` child")
        return self._dialog_model.createInstance(ctl.__ooo_full_ns__)

    def add_control(self, name: str, ctl: 'UnoControlDialogElement') -> ControlWrapper:
        if name.strip() == "":
            raise ValueError(
                f"{self.__class__.__name__}.add_control() 'ctl.Name' cannot be an empty string.")
        if name in self._ctl_names:
            raise Exception(
                "Control by the name of '{0}' already exist".format(name))
        self._dialog_model.insertByName(name, ctl)
        self._ctl_names.add(name)
        control: 'XWindow' = self._dialog.getControl(name)
        control.setPosSize(ctl.PositionX, ctl.PositionY, ctl.Width,
                           ctl.Height, PosSize.POSSIZE)
        c_wrap = ControlWrapper(name, ctl, ctl.PositionX, ctl.PositionY)
        return c_wrap

    def _create(self, name: str) -> 'XInterface':
        return self._ctx.getServiceManager().createInstanceWithContext(name, self._ctx)

    # region Properties

    @property
    def control_count(self) -> int:
        """
        Get the number of control in instance of class

        Returns:
            int: number of controls  or 0 if there are no current controls
        """
        return len(self._ctl_names)

    @property
    def dialog(self) -> 'UnoControlDialog':
        ''' Gets current Dialog
        '''
        return self._dialog

    @property
    def model(self) -> 'UnoControlDialogModel':
        ''' Gets current Dialog Model
        '''
        return self._dialog_model
    # endregion Properties
