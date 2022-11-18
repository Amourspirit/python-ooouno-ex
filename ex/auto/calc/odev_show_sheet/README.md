# Show Sheet

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/194169727-f5a61ab2-e336-42c3-8ef1-31299b81100d.jpg" width="348" height="232">
</p>

Example of opening a spreadsheet and inputting a password (`foobar`) to unlock sheet.

Also demonstrates how to create input password dialog and message dialog.

Optionally saves the input file as a new file.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

## Automate

A message box is display once the document has been processed asking if you want to close the document.

The following command will run automation that opens Calc document and ask for password.

### Cross Platform

From this folder.

```sh
python -m start --show --file "../../../../resources/ods/totals.ods"
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_show_sheet/start.py --show --file "resources/ods/totals.ods" --out "tmp/totals.pdf"
```

Alternatively

```sh
python ./ex/auto/calc/odev_show_sheet/start.py --show --file "resources\data\sorted.csv" --out "tmp/totals.html"
```

### Windows

```ps
python .\ex\auto\calc\odev_show_sheet\start.py --show --file "resources\ods\totals.ods" --out "tmp/totals.pdf"
```

Alternatively

```ps
python .\ex\auto\calc\odev_show_sheet\start.py --show --file "resources\data\sorted.csv" --out "tmp/totals.html"
```
