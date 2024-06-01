# Writer Custom Properties

This is a simple example that show how to add custom properties to a Writer document. The custom properties are stored in the document and can be accessed later.

This example uses [OOO Development Tools] (OooDev) which makes custom properties possible.

In some cases it may be beneficial to add custom properties to a document. This can be used to store meta data or other information that is not part of the document content.

[OOO Development Tools] version `0.45.0` or higher is required.

## See

See Also:

- [OOO Development Tools]
- [Draw Custom Page Properties](../../draw/odev_custom_page_props#readme)
- [Impress Custom Page Properties](../../impress/odev_custom_page_props#readme)
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
python ./ex/auto/writer/odev_custom_props/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_custom_props\start.py
```


### Output

```text
Meta1: Some meta data
Meta2: Some other meta data
Prop1: Hello
Prop2: World
Prop3: 777
Saving document with custom properties
File: /tmp/props.odt

Opened document again. Custom properties:
Meta1: Some meta data
Meta2: Some other meta data
Prop1: Hello
Prop2: World
Prop3: 777.0
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
