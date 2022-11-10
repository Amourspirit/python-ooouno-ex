# Impress read text data and convert to slides with points

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/200890963-b4569f69-a647-465a-9154-0ec114c45121.png" width="800" height="406">
</p>

Convert a text file of points into a series of slides. Uses a template from Office.

A message box is display once the document has been created asking if you want to close the document.

This demo uses [OOO Development Tools]

See Also:

- [OOO Development Tools - Part 3: Draw & Impress](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/index.html)

## Automate

A single parameters can be passed in which is the slide show document to data read from:

**Example:**

```sh
python ./ex/auto/impress/odev_points_builder/start.py "resources/data/pointsInfo.txt"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_points_builder/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_points_builder\start.py
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/