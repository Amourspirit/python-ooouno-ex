# Slide Chart

![slide_chart](https://user-images.githubusercontent.com/4193389/198894178-1c6b79bf-185f-44e0-b061-3c026da88384.png)

Generates a column chart using the "Sneakers Sold this Month" table from `chartsData.ods`, copies it to the clipboard.
Then an Impress document is created, and the chart image is pasted into it.
The Chart is then moved to the center of the slide.

Also demonstrates saving the chart as an image and working with multiple documents.

A message box is display once the document has been created asking if you want to close the documents.

## NOTE

There is currently a [bug](https://bugs.documentfoundation.org/show_bug.cgi?id=151846) in LibreOffice `7.4` that does not allow the `Chart2` class to load.
The `Chart2` has been tested with LibreOffice `7.3`. The bug has been fixed in `7.5`.

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
python ./ex/auto/chart2/slide_chart/start.py
```

### Windows

```ps
python .\ex\auto\chart2\slide_chart\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/slide_chart'
```

This will copy the `slide_chart` example to the examples folder.

In the terminal run:

```bash
cd slide_chart
python -m start
```
