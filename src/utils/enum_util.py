# coding: utf-8
# see: https://tinyurl.com/yc67tywv

def enum_class_new_flex(cls, value):
    """
    New (__new__) method enum classes that are created from attrib value

    Args:
        value (object): Can be Enum, Enum.value, str

    Raises:
        ValueError: if unable to match enum instance

    Returns:
        (enum): of enum subclass

    Example:
        ..code-block:: python
        
            >>> class HorizontalAlignment(IntEnum):
            >>>     RIGHT = auto()
            >>>     LEFT = auto()
            >>>     CENTER = auto()
            >>>
            >>> e = HorizontalAlignment("RIGHT")
            >>> print(e.value)
            RIGHT
            >>> e = HorizontalAlignment(HorizontalAlignment.LEFT)
            >>> print(e.value)
            LEFT
            >>> e = HorizontalAlignment(HorizontalAlignment.CENTER.value)
            >>> print(e.value)
            CENTER
    """
    if isinstance(value, str):
        if hasattr(cls, value):
            return getattr(cls, value)
    _type = type(value)
    if _type is cls:
        return value
    raise ValueError("%r is not a valid %s" % (value, cls.__name__))
