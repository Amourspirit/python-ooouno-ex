# Impress Modify Pages

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198418648-34ab1937-9d3d-4a10-bf38-54f4abc39775.png">
</p>

Add two new slides to the input document

- add a title-only slide with a graphic at the end
- add a title/subtitle slide at the start

A message box is display once the slide show has ended asking if you want to close the document.

This demo uses [OOO Development Tools]

## Automate

A single parameters can be passed in which is the slide show document to modify:

**Example:**

```sh
python ./ex/auto/impress/odev_modify_slides/start.py "resources/presentation/algsSmall.ppt"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_modify_slides/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_modify_slides\start.py
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/