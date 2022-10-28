# Impress Auto Slide Show

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198406431-6b28b28b-4949-4a41-bf67-ff485ab964a2.png" width="659" height="448">
</p>

Demonstrates displaying a slide show that automatically plays using default `algs.odp` file.

A message box is display once the slide show has ended asking if you want to close the document.

This demo uses [OOO Development Tools]

## Automate

A single parameters can be passed in which is the slide show document to modify:

**Example:**

```sh
python ./ex/auto/impress/odev_auto_show/start.py "resources/presentation/algs.ppt"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_auto_show/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_auto_show\start.py
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/