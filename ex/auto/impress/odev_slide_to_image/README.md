# Impress Slide to Image

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198423388-f8845bec-781a-42ef-b8cf-20bb13b9cb43.png">
</p>

Saves a given page of a slide presentation (e.g. ppt, odp) as an image file (e.g. "gif", "png", "jpeg", "wmf", "bmp", "svg")

This demo uses [OOO Development Tools]
## Automate

Args are required to be passed.

**Get Help:**

```sh
python -m start --help
```

If `--output_dir` arg is omitted then a temporary dir is created in the systems temp dir and the output image is saved there.

### Cross Platform

From current example folder.

```sh
python -m start --file "../../../../resources/presentation/algs.ppt" --out_fmt "jpeg" --idx 0
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_slide_to_image/start.py --file "resources/presentation/algs.ppt" --out_fmt "jpeg" --idx 0
```

### Windows

```ps
python .\ex\auto\impress\odev_slide_to_image\start.py --file "resources/presentation/algs.ppt" --out_fmt "jpeg" --idx 0
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/