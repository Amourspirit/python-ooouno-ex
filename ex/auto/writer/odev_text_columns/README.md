# Text Columns Example

<p align="center">
  <img width="685" height="404" src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/14d208b5-0cca-437a-aeb7-9ca1a57c97f7">
</p>

This example demonstrates how to create a text document with columns using the [OOO Development Tools].
The original is in Java and can be found at [LibreOffice Developer's Guide: Chapter 7 - Text Documents](https://wiki.documentfoundation.org/Documentation/DevGuide/Text_Documents#Columns).

Text frames, text sections and page styles can be formatted to have columns. The width of columns is relative since the absolute width of the object is unknown in the model. The layout formatting is responsible for calculating the actual widths of the columns.

Columns are applied using the property TextColumns. It expects a `com.sun.star.text.TextColumns` service that has to be created by the document factory. The interface `com.sun.star.text.XTextColumns` refines the characteristics of the text columns before applying the created TextColumns service to the property TextColumns.


## Automate

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

From project root folder.

```sh
python ./ex/auto/writer/odev_text_columns/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_text_columns\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_text_columns'
```

This will copy the `odev_text_columns` example to the examples folder.

In the terminal run:

```bash
cd odev_text_columns
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
