# Calc Sheets Custom Properties

This is a simple example that show how to add custom properties to Calc Sheets. The custom properties are stored in the document sheets and can be accessed later.

This example uses [OOO Development Tools] (OooDev) which makes custom properties possible.

In some cases it may be beneficial to add custom properties to a Sheet. This can be used to store meta data or other information that is not part of the sheet content.

[OOO Development Tools] version `0.45.0` or higher is required.

## See

See Also:

- [OOO Development Tools]
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


### Output

```text
Custom properties for Sheet1:
Meta1: Some meta data
Meta2: Some other meta data
Prop1: Hello
Prop2: World
Prop3: 777
Custom properties for Sheet2:
GoalYear: 2030
SheetOwner: Elon Musk
SheetPurpose: Save a planet
Saving document with custom properties
File: /tmp/props.ods

Opened document again. Custom properties:
Custom properties for Sheet2:
GoalYear: 2030
SheetOwner: Elon Musk
SheetPurpose: Save a planet
Custom properties for Sheet2:
GoalYear: 2030.0
SheetOwner: Elon Musk
SheetPurpose: Save a planet
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
