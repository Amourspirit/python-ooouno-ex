# Calc Cell Custom Properties

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/403a7be8-96d0-4350-82ab-56b9cafb8c46" width="366" height="479">
</p>


This is a example that show how to add custom properties to Calc sheet cells. The custom properties are stored in the document and can be accessed later. Using `OooDev` custom properties can be added to cells in a Calc document. The can store values that are not part of the cell content such as booleans, string and numbers.

See also [ooodev.calc.CalcCell](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/calc/calc_cell.html).

Cell custom property method:

```python
cell = sheet["A1"]
cell.set_custom_property("OriginalValue", cell.value)
cell.set_custom_property("IsOriginal", True)
assert cell.has_custom_properties()
assert cell.has_custom_property("OriginalValue")
cell.remove_custom_property("OriginalValue")
cell.remove_custom_properties() # removes all custom properties
```

This example uses [OOO Development Tools] (OooDev) which makes custom properties possible.

This example also demonstrates how to intercept context menu to add custom actions for the sheet cells.
for more on this see [Intercept Context Menu] example.

In some cases it may be beneficial to add custom properties to a sheet cells. This can be used to store meta data or other information that is not part of the sheet content.
This may be extra helpful when using `OooDev` to help build Excel Addons.

In this example a Calc document is generated using the [generate_data.py] script. The [generate_data.py] script creates a Calc document with a sheet that has a table of data.
The script then adds custom properties to some of the cells in the table. The custom properties are stored in the document and can be accessed later.

In the `generate_data.py` script.

```python

def set_original_prop(sheet: CalcSheet) -> None:
    rng = sheet.rng("B2:C21")
    for cell_obj in rng:
        cell = sheet[cell_obj]
        cell.set_custom_property("OriginalValue", cell.value)
```

Two new custom menu items are added to the context menu via an interception. The menu items are only injected if the current cell value does not match the custom property `OriginalValue` value. The two menu items are `Reset to Original` and `Update Original`. The `Reset to Original` menu item will reset the cell value to the `OriginalValue` custom property value. The `Update Original` menu item will update the `OriginalValue` custom property value to the current cell value.

This is for demonstration purposes.


The `odev_custom_cell_props_lib` folder contains the dispatch classes that are used to intercept the menu and handle the menu items.


Cell custom properties are not part of the LibreOffice Calc API. They are a feature of the `OooDev` library.
For this reason, some small clean is recommended when using custom properties. This is not required as in most cases `OooDev` will clean up after itself for copied cells and some other operation. When a cell is deleted `OooDev` will not clean up any custom properties for the deleted cell automatically. This is because `OooDev` does not attach listeners to the cells or sheets for custom properties.
The `ooodev.calc.cell.custom_prop_clean.CustomPropClean` class can be used to clean up custom properties for a sheet or cell. The `clean` method can be called to clean up custom properties for a cell or sheet.
In this example the clean method is called automatically when the document is saving.
If no cleanup is done, it will not have any ill effects on the document. The custom properties will just be left in the document.

## Building the Example

In this folder in the terminal run the following commands.

The `data/custom_props.ods` is already generated and included in the repo. If you want to generate the document yourself run the following commands.


```sh
make data
make build
```

`make data` will generate the data soruce file for the Calc document. Running this command will generate an error when Calc is closing but this is expected.
`make build` will build embed the python source code into the document from the previous step an save it as `data/custom_props.ods`.

The `data/custom_props.ods` file will be created with the custom properties added to the cells. The document requires the `OooDev` extension to be install to run the custom properties code.

Alternatively, you can run the `start.py` script in a code editor with a break point to see the custom properties code in action.


[OOO Development Tools] version `0.45.0` or higher is required.

## See Also


- [OOO Development Tools]
- [Intercept Context Menu]
- [Calc Custom Properties](../odev_custom_sheet_props#readme)
- [Draw Custom Page Properties](../../draw/odev_custom_page_props#readme)
- [Impress Custom Page Properties](../../impress/odev_custom_page_props#readme)
- [Writer Custom Properties](../../writer/odev_custom_props#readme)

See [source code](./start.py)

## Automate

### Dev Container

From project root folder.

```sh
python -m start
```

### Cross Platform

From project root folder.

```sh
python -m start
```

### Linux/Mac

Run from current example folder.

```sh
python ./ex/auto/calc/odev_custom_props/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\calc\odev_custom_props\start.py
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[Intercept Context Menu]: ../odev_context_link#readme
[generate_data.py]: ./generate_data.py