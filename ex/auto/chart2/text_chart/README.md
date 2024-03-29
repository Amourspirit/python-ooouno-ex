# Text Chart

![text_chart](https://user-images.githubusercontent.com/4193389/198876133-15803e66-008c-4eeb-a2ae-28021a0e7245.png)

Generates a column chart using the "Sneakers Sold this Month" table from `chartsData.ods`, copies it to the clipboard, and closes the spreadsheet. Then a text document is created, and the chart image is pasted into it.

A message box is display once the document has been created asking if you want to close the document.

## NOTE

There is currently a [bug](https://bugs.documentfoundation.org/show_bug.cgi?id=151846) in LibreOffice `7.4` that does not allow the `Chart2` class to load.
The `Chart2` has been tested with LibreOffice `7.3`

## Automate

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
python ./ex/auto/chart2/text_chart/start.py
```

### Windows

```ps
python .\ex\auto\chart2\text_chart\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/text_chart'
```

This will copy the `text_chart` example to the examples folder.

In the terminal run:

```bash
cd text_chart
python -m start
```
