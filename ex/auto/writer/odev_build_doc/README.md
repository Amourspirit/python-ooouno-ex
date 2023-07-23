# Build Doc

Create a new Writer document, add a few lines, styles,
images, text frame, bookmark, and displays a message box giving the the option to save document
as the file "build.odt" in the current working folder.

## See

See Also:

- [Text Styles]
- [Style Changes to Words and Phrases]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Cross Platform

From this folder.

```shell
python -m start
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/writer/odev_build_doc/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_build_doc\start.py
```

![build_doc](https://user-images.githubusercontent.com/4193389/184692062-4554d35d-4be8-4aac-99a6-4d7962e2017b.gif)

## Output

[build.odt](./data/build.odt)

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_doc'
```

This will copy the `odev_build_doc` example to the examples folder.

In the terminal run:

```bash
cd odev_build_doc
python -m start
```

[Text Styles]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter06.html
[Style Changes to Words and Phrases]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter06.html#style-changes-to-words-and-phrases
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
