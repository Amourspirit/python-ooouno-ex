# Rotate Shape Macro

Documents a macro for rotating a shape in a draw document.

The [macro/rotate_shape/rotate_shape.py](../../../macro/draw/rotate_shape.py) file contains the macro code.

This macro rotates a shape in a draw document.

This macro is a simple example of how to rotate a shape in a draw document and How to get user input.



## Running Macro

Start a draw document and insert a shape. Then run the macro.

On Tools menu select `Macros > Run Macro`. In the dialog that appears, navigate to the location of the macro file and select it.

![Macro Selection Box](https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/a7cfadf0-a77f-400a-81b8-398736292223)

Click the ``Run`` button.

A dialog will appear asking for the angle to rotate the shape.

![Rotation Angle Dialog Options](https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/ccdd951b-7248-48cf-86a2-f8dba550aa65)

The shapes will rotate the specified angle.

![Screen shot of rotated images](https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/d8232d0d-c369-4507-87dd-5ae1d333fcb9)

## Running Macro in Codespace

This examples here can be run in a codespace (Development container).
It is beyond the scope of this document to explain how to run the codespace;
however, [Live LibreOffice Python](https://github.com/Amourspirit/live-libreoffice-python) is very similar and has a [wiki](https://github.com/Amourspirit/live-libreoffice-python/wiki) to get you started.

## Running Macro Locally

This examples here can be run locally like any other python macro.

This example requires [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html), version `0.34.0` or later, which can also be installed as an extension in LibreOffice. See [OOO Development Tools](https://extensions.libreoffice.org/en/extensions/show/41700) Extension.