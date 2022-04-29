# coding: utf-8
import uno
import string
from typing import TYPE_CHECKING, Optional, Tuple
from ooo.dyn.beans.property_value import PropertyValue
from ooo.dyn.beans.method_concept import MethodConcept
from ooo.helper import uno_helper as ooo_uno_helper
from ooo.dyn.frame.the_desktop import theDesktop
from ooo.dyn.beans.the_introspection import theIntrospection
from ooo.dyn.reflection.the_core_reflection import theCoreReflection
import scriptforge as SF

if TYPE_CHECKING:
    from ooo.lo.drawing.drawing_document import DrawingDocument
    from ooo.lo.lang.service_manager import ServiceManager
    from ooo.lo.lang.x_component import XComponent
    from ooo.lo.lang.x_multi_component_factory import XMultiComponentFactory
    from ooo.lo.reflection.x_idl_method import XIdlMethod
    from ooo.lo.style.style import Style
    from ooo.lo.style.style_family import StyleFamily
    from ooo.lo.uno.x_component_context import XComponentContext
    from ooo.lo.uno.x_interface import XInterface
    from ooo.lo.frame.x_frame import XFrame
    from ooo.lo.frame.x_model import XModel


# region Cached Values
# It is cached in a global variable.
_star_desktop: theDesktop = None
# _go_service_manager the ServiceManager of the running OOo.
# It is cached in a global variable.
# _go_service_manager: Optional[XMultiComponentFactory] = None
_go_service_manager: "XMultiComponentFactory" = None

# The theCoreReflection object is cached in a global variable.
_the_core_reflection: theCoreReflection = False

_star_xmodel = None


def reset_cached():
    """Resets globals. Use in testing scenarios."""
    global _star_desktop
    global _go_service_manager
    global _the_core_reflection
    global _star_xmodel
    _star_desktop = None
    _go_service_manager = None
    _the_core_reflection = False
    _star_xmodel = None


# endregion Cached Values


def set_global(**kwargs):
    """
    Set global vars for module

    Keyword Args:
        desktop (object): Sets the desktop used in the get_desktop() function.
        sm (XMultiComponentFactory): Sets the service manager that is used for creating instances of objects.

    Note:
        This method is used primarly in testing
    """
    global _star_desktop, _go_service_manager, _the_core_reflection
    if "desktop" in kwargs:
        _star_desktop = kwargs["desktop"]
    if "sm" in kwargs:
        _go_service_manager = kwargs["sm"]


def get_service_manager() -> "ServiceManager":
    """
    Get the ServiceManager from the running OpenOffice.org.
    Then retain it in the global variable go_service_manager for future use.
    This is similar to the GetProcessServiceManager() in OOo Basic.

    Returns:
        XMultiComponentFactory: Service Manager
    """
    global _go_service_manager
    if not _go_service_manager:
        # Get the uno component context from the PyUNO runtime
        ctx: "XComponentContext" = uno.getComponentContext()
        _go_service_manager = ctx.getServiceManager()
    return _go_service_manager


def create_uno_service(
    clazz: str,
    ctx: "Optional[XComponentContext]" = None,
    args: Optional[Tuple[object]] = None,
) -> "XInterface":
    """A handy way to create a global objects within the running OOo.
    Similar to the function of the same name in OOo Basic.

    Args:
        clazz (str, object): name of the service to be instanciated or object that contains __ooo_full_ns__ attribute.
            If ``clazz`` is class then return service has ``__ooo_ns__``, ``__ooo_full_ns__`` and ``__ooo_type_name__`` atribute values set.
        ctx (XComponentContext, optional): the context if required.
        args (Typle[object], optional): the arguments when needed.

    Returns:
        object: component instance

    Notes:
        A service signals that it expects parameters during instantiation by supporting the com.sun.star.lang.XInitialization interface.
        There maybe services which can only be instantiated with parameters
    """
    return ooo_uno_helper.create_uno_service(clazz=clazz, ctx=ctx)


def get_desktop() -> theDesktop:
    """An easy way to obtain the Desktop object from a running OOo.

    Returns:
        object: com.sun.star.frame.Desktop

    See Also:
        `LibreOffice 7.2 SDK API Desktop Reference <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1Desktop.html>`_

    Warning:
        Method uses "com.sun.star.frame.Desktop" and the
        `Desktop API <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1Desktop.html>`_
        states that it is deprecated.
        The api suggest; Rather use the 'theDesktop' singleton.

    Todo:
        Test if this method would work with 'theDesktop'
    """
    global _star_desktop
    if _star_desktop == None:
        bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
        # StarDesktop = create_uno_service("com.sun.star.frame.Desktop")
        _star_desktop = bas.StarDesktop
    return _star_desktop


# preload the StarDesktop variable.
# get_desktop()


def get_core_reflection() -> theCoreReflection:
    """
    This service is the implementation of the reflection API.

    Returns:
        object: com.sun.star.reflection.CoreReflection

    See Also:
        `LibreOffice 7.2 SDK CoreReflection API Reference <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1reflection_1_1CoreReflection.html>`_

    Warning:
        Method uses "com.sun.star.reflection.CoreReflection" and the
        `CoreReflection API <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1reflection_1_1CoreReflection.html>`_
        states that it is deprecated.
        The api suggest; Rather use the 'theCoreReflection' singleton.

    Todo:
        Test if this method would work with 'theCoreReflection'
    """
    # https://stackoverflow.com/questions/67527942/why-cant-an-annotated-variable-be-global
    global _the_core_reflection
    if not _the_core_reflection:
        _the_core_reflection = theCoreReflection()
        # go_core_reflection = create_uno_service(
        #     "com.sun.star.reflection.CoreReflection")
    return _the_core_reflection


def create_uno_struct(cTypeName: str) -> object:
    """Create a UNO struct and return it.
    Similar to the function of the same name in OOo Basic.

    Returns:
        object: uno struct
    """
    oCoreReflection = get_core_reflection()
    # Get the IDL class for the type name
    oXIdlClass = oCoreReflection.forName(cTypeName)
    # Create the struct.
    oReturnValue, oStruct = oXIdlClass.createObject(None)
    return oStruct


def get_xframe() -> "XFrame":
    """
    Gets xFrame for current LibreOffice instance

    Returns:
        XFrame: xFrame
    """
    desk_top = get_desktop()
    frame = desk_top.getCurrentFrame()
    return frame


def get_xModel() -> "XModel":
    """
    Gets xModel

    Returns:
        object: Gets xModel for current LibreOffice instance
    """
    global _star_xmodel
    if _star_xmodel == None:
        bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
        bas.ThisComponent.CurrentController.Frame
        # xFrame = get_xframe()
        _star_xmodel = bas.ThisComponent.CurrentController.Frame.getController().getModel()
    return _star_xmodel


# ------------------------------------------------------------
#   API helpers
# ------------------------------------------------------------

# region API helpers
def has_uno_interface(oObject: object, cInterfaceName: str) -> bool:
    """
    Gets if if :paramref:`~.has_uno_interface.oObject` implements :paramref:`~.has_uno_interface.cInterfaceName`

    Args:
        oObject (object): Object to test for interface.
        cInterfaceName (str): Name of intefact to test for.

    Returns:
        bool: True if :paramref:`~.has_uno_interface.oObject` implements :paramref:`~.has_uno_interface.cInterfaceName`
        ; Otherwise, False

    Note:
        Similar to Basic's has_uno_interfaces() function, but singular not plural.
    """

    # Get the Introspection service.
    # oIntrospection: 'Introspection' = create_uno_service(
    #     "com.sun.star.beans.Introspection")
    oIntrospection = theIntrospection()

    # Now inspect the object to learn about it.
    oObjInfo = oIntrospection.inspect(oObject)

    # Obtain an array describing all methods of the object.
    oMethods: "Tuple[XIdlMethod]" = oObjInfo.getMethods(MethodConcept.ALL)
    # Now look at every method.
    for oMethod in oMethods:
        # Check the method's interface to see if
        #  these aren't the droids you're looking for.
        cMethodInterfaceName: str = oMethod.getDeclaringClass().getName()
        if cMethodInterfaceName == cInterfaceName:
            return True
    return False


def has_uno_interfaces(oObject: "XInterface", *cInterfaces: "XInterface") -> bool:
    """
    Gets if :paramref:`~.has_uno_interfaces.oObject` implements all :paramref:`~.has_uno_interfaces.cInterfaces`

    Args:
        oObject (object): object to test for interface(s)
        cInterfaces (str): one or more interfaces to test

    Returns:
        bool: ``True`` if :paramref:`~.has_uno_interfaces.oObject` implements all :paramref:`~.has_uno_interfaces.cInterfaces`
        ; Otherwise, ``False``

    Note:
        Similar to the function of the same name in OOo Basic.
    """
    for cInterface in cInterfaces:
        if not has_uno_interface(oObject, cInterface):
            return False
    return True


# endregion API helpers

# region Styles
def define_style(
    oDrawDoc: "DrawingDocument",
    cStyleFamily: str,
    cStyleName: str,
    cParentStyleName: Optional[str] = None,
) -> "Style":
    """Add a new style to the style catalog if it is not already present.
    This returns the style object so that you can alter its properties.
    """
    # oStyleFamily = oDrawDoc.get_styleFamilies().getByName(cStyleFamily)
    oStyleFamily: "StyleFamily" = oDrawDoc.getStyleFamilies().getByName(cStyleFamily)

    # Does the style already exist?
    if oStyleFamily.hasByName(cStyleName):
        # then get it so we can return it.
        oStyle: "Style" = oStyleFamily.getByName(cStyleName)
    else:
        # Create new style object.
        oStyle: "Style" = oDrawDoc.createInstance("com.sun.star.style.Style")

        # Set its parent style
        if cParentStyleName != None:
            oStyle.setParentStyle(cParentStyleName)

        # Add the new style to the style family.
        oStyleFamily.insertByName(cStyleName, oStyle)

    return oStyle


def get_style(
    oDrawDoc: "DrawingDocument", cStyleFamily: str, cStyleName: str
) -> "Style":
    """Lookup and return a style from the document."""
    return oDrawDoc.getStyleFamilies().getByName(cStyleFamily).getByName(cStyleName)


# endregion Styles

# region General Utility functions
def convert_to_url(cPathname: str) -> str:
    """
    Convert a Windows or Linux pathname into an OOo URL.

    Args:
        cPathname (str): Path Name to Convert

    Returns:
        string: Converted path to url.
    """
    if len(cPathname) > 1:
        if cPathname[1:2] == ":":
            cPathname = "/" + cPathname[0] + "|" + cPathname[2:]
    cPathname = string.replace(cPathname, "\\", "/")
    cPathname = "file://" + cPathname
    return cPathname


def load_component_from_url(
    cUrl: str, tProperties: "Tuple[PropertyValue]"
) -> "XComponent":
    """
    Open or Create a document from it's URL.

    Args:
        cUrl (str): specifies the URL of the document to load
            New documents are created from URL's such as:

            * ``private:factory/sdraw``
            * ``private:factory/swriter``
            * ``private:factory/scalc``
            * ``private:factory/simpress``

        tProperties (:obj:`com::sun::star::beans::PropertyValue`, optional): Properties

    Returns:
        com::sun::star::lang::XComponent: ``XComponent`` for successfully loaded documents or
        ``None`` if it failed.

    See Also:
        `LibreOffice API - XComponentLoader Interface Reference <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html>`_
    """
    StarDesktop = get_desktop()
    oDocument = StarDesktop.loadComponentFromURL(cUrl, "_blank", 0, tProperties)
    return oDocument


# endregion General Utility functions


# lists the scripts, that shall be visible inside OOo. Can be omitted, if
# all functions shall be visible
g_exportedScripts = ()
