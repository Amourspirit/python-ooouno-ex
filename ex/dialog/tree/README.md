# Tab and Tree Control Dialog Example

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/284018833-91fdd4ac-a2c2-4105-a32b-922480240a12.png" width="498" height="433">
</p>

This example demonstrates how to create a Dialog window programmatically adding Tabs and Tree controls and text edit controls to the dialog.
Searching is implemented and the tree nodes can be edited when `Allow Tree Node Editing` is checked.
Normal node editing is implemented by pressing `F2` on a node or double clicking on it to enter in to edit mode.

Each tab in the dialog is a separate class the implements the logic for that tab.

This demo uses This demo uses [OOO Development Tools] (OooDev).
Also available as a LibreOffice [Extension](https://extensions.libreoffice.org/en/extensions/show/41700).

See Also:

- [ooodev.dialog.dialogs.Dialogs](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dialogs.html)
- [ooodev.dialog.dl_control.ctl_dialog.CtlDialog](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_dialog.html)
- [ooodev.dialog.dl_control.ctl_tab_page.CtlTabPage](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_tab_page.html)
- [ooodev.dialog.dl_control.ctl_tree.CtlTree](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_tree.html)
- [ooodev.dialog.dl_control.ctl_button.CtlButton](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_button.html)
- [ooodev.dialog.dl_control.ctl_check_box.CtlCheckBox](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/dl_control/ctl_check_box.html)

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
python ./ex/dialog/tree/start.py
```

### Windows

```ps
python .\ex\dialog\tree\start.py
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

After running the build command, the output will be in the `build/tree_tabs` folder of this projects root.
The file name will be `tabs_tree.odt`.

See [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html).

### Running the embedded Dialog

To run the embedded dialog, open the `tabs_tree.odt` file in LibreOffice and run
`Tools -> Macros -> Run Macro...` and select `tabs_tree.ods -> tabs_tree -> show_tabs` and click `Run`.

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/tree'
```

This will copy the `gird` example to the examples folder.

In the terminal run:

```bash
cd tree
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OOO Development Tools Extension]: https://extensions.libreoffice.org/en/extensions/show/41700
