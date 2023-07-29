# Impress Slide to Image

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198423388-f8845bec-781a-42ef-b8cf-20bb13b9cb43.png">
</p>

Saves a given page of a slide presentation (e.g. ppt, odp) as an image file (e.g. "gif", "png", "jpeg", "wmf", "bmp", "svg")

This demo uses [OOO Development Tools]

See Also:

- [OOO Development Tools - Chapter 17. Slide Deck Manipulation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/chapter17.html)

## Automate

Args are required to be passed.

**Get Help:**

```sh
python -m start --help
```

If `--output_dir` arg is omitted then a temporary dir is created in the systems temp dir and the output image is saved there.

### Dev Container

From current example folder.

```sh
python -m start --file "./data/algs.ppt" --out_fmt "jpeg" --idx 0
```

### Cross Platform

From current example folder.

```sh
python -m start --file "./data/algs.ppt" --out_fmt "jpeg" --idx 0
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_slide_to_image/start.py --file "ex/auto/impress/odev_slide_to_image/data/algs.ppt" --out_fmt "jpeg" --idx 0
```

### Windows

```ps
    python .\ex\auto\impress\odev_slide_to_image\start.py --file "ex/auto/impress/odev_slide_to_image/data/algs.ppt" --out_fmt "jpeg" --idx 0
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_slide_to_image'
```

This will copy the `odev_slide_to_image` example to the examples folder.

In the terminal run:

```bash
cd odev_slide_to_image
python -m start -h
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
