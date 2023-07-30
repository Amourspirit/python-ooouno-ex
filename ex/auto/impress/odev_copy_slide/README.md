# Impress copy Slide

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198409975-eb50eb6a-3216-4cbb-a315-2ebc4a5d9bf7.png" width="253" height="313">
</p>

Demonstrates opening a presentation file in Impress and copying a slide from a given index to after another given index.

A message box is display once the document has been created asking if you want to close the document.

This demo uses [OOO Development Tools]

See Also:

- [OOO Development Tools - Chapter 17. Slide Deck Manipulation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/chapter17.html)

## Automate

Three parameters can be passed in:

1. File Name: Such as `"resources/presentation/algs.odp"`
2. From Index: Such as `2`
3. To Index: Such as `4`

**Example:**

```sh
python ./ex/auto/impress/odev_copy_slide/start.py "resources/presentation/algs.odp" 0 2
```

If no parameters are passed then the script is run with the above parameters.

### Dev Container

From current example folder.

```sh
python -m start
```

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_copy_slide/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_copy_slide\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_copy_slide'
```

This will copy the `odev_copy_slide` example to the examples folder.

In the terminal run:

```bash
cd odev_copy_slide
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
