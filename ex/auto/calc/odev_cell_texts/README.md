# Cell Texts

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/204042072-637c86d5-5045-44e4-8cac-f499f73a8dec.png">
</p>

Demonstrates the following:

- Add paragraphs, hyperlink to a cell.
- Change the style of the cells.
- Access the text.
- Add an annotation.

This demo uses This demo uses [OOO Development Tools] (ODEV).

See Also:

- [OOO Development Tools - Chapter 24. Complex Data Manipulation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter24.html)

## Automate

A message box is display once the document has been processed asking if you want to close the document.

### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_cell_texts/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_cell_texts\start.py
```

## Example Usage

Starts LibreOffice Calc, runs automation and saves output to `tmp/linkedText.ods`

```ps
python .\ex\auto\calc\odev_cell_texts\start.py -o "tmp/linkedText.ods"
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/