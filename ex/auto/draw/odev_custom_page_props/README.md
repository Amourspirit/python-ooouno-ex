# Draw Custom Page Properties

This is a simple example that show how to add custom properties to a Draw document Pages. The custom properties are stored in the draw pages and can be accessed later.

This example uses [OOO Development Tools] (OooDev) which makes custom properties possible.

In some cases it may be beneficial to add custom properties to a Draw pages. This can be used to store meta data or other information that is not part of the document content.

[OOO Development Tools] version `0.45.0` or higher is required.

## See

See Also:

- [OOO Development Tools]
- [Impress Custom Page Properties](../../impress/odev_custom_page_props#readme)
- [Writer Custom Properties](../../writer/odev_custom_props#readme)
- [Calc Custom Properties](../../calc/odev_custom_sheet_props#readme)
- [Custom Cell Properties](../../calc/odev_custom_cell_props#readme)

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
python ./ex/auto/draw/odev_custom_page_props/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\draw\odev_custom_page_props\start.py
```


### Output

```text
Meta1: Some meta data
Meta2: Some other meta data
Prop1: Hello
Prop2: World
Prop3: 777
Saving document with custom properties
File: /tmp/props.odg

Opened document again. Custom properties:
Meta1: Some meta data
Meta2: Some other meta data
Prop1: Hello
Prop2: World
Prop3: 777.0
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
