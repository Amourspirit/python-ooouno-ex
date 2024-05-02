# Radio Button2 Example

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/283243452-94e5910a-86fb-4d45-ad47-2cb21b266ac4.png" width="335" height="277">
</p>

This example demonstrates how to create a Dialog window programmatically and add a Radio Button Groups.

This demo is functionally the same as the [Radio Button](../radio_groupbox/) Example. The difference is that this example uses the `CalcDoc` class to create an instance of the `Dialog` class that build the dialog box and its controls.

By using the built in `create_dialog()` method of the `CalcDoc` class it ensure that the dialog and its controls get the correct context and parent window.
This can be important in a multi document environment.

This demo uses This demo uses [OOO Development Tools] (OooDev).
Also available as a LibreOffice [Extension](https://extensions.libreoffice.org/en/extensions/show/41700).

See Also:

- [ooodev.dialog.dialogs.Dialogs](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dialogs.html)
- [ooodev.dialog.dl_control.ctl_dialog.CtlDialog](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_dialog.html)
- [ooodev.dialog.dl_control.ctl_radio_button.CtlRadioButton](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_radio_button.html)
- [ooodev.dialog.dl_control.ctl_button.CtlButton](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_button.html)

### Dev Container

From this folder.

```sh
python -m start
```

### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/dialog/radio_groupbox2/start.py
```

### Windows

```ps
python .\ex\dialog\grid\start.py
```

## Embedding

### Embedding the Dialog in a document

The dialog can be embedded in a Writer document in the following two ways:

The recommended way is to have [OOO Development Tools Extension] installed and use the number `1` method below.

These commands must be run from this current folder in a terminal.

1. **Embedding the Dialog in a Writer document for use with extension**

   The dialog can be embedded in Writer document running the following command in the terminal:

   ```sh
   make build
   ```

   With extension installed this will create a much more lightweight file. The startup for the macro will be faster.

2. **Embedding the Dialog in a Writer document for use without extension**

    With this method the macro will run without the extension installed.

   The dialog can be embedded in a Writer document running the following command in the terminal:

   ```sh
   make build_ooodev
   ```

    This will include the required `ooodev` packages in the document. This will create a much larger file. The startup for the macro will be slower. After first load the macro files will be cached and running the macro again will be faster.

### Build output

After running the build command, the output will be in the `build/radio_example2` folder of this projects root.
The file name will be `radio_example.odt`.

See [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html).

### Running the embedded Dialog

To run the embedded dialog, open the `tabs_tree.odt` file in LibreOffice and run
`Tools -> Macros -> Run Macro...` and select `radio_example.ods -> radio_example -> show_radio_example` and click `Run`.

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/radio_groupbox2'
```

This will copy the `radio_groupbox2` example to the examples folder.

In the terminal run:

```bash
cd radio_groupbox2
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OOO Development Tools Extension]: https://extensions.libreoffice.org/en/extensions/show/41700
