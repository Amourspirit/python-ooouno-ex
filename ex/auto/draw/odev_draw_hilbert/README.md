# Draw Hilbert curve

![img](https://upload.wikimedia.org/wikipedia/commons/0/06/Hilbert_curve_3.svg)

Generate a [Hilbert curve] of the specified level.

Created using a series of rounded blue lines.
Position/size the window, resize the page view

This demo uses [OOO Development Tools]

See Also:

- [OOO Development Tools - Part 3: Draw & Impress](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/index.html)

## Usage

run `python -m draw_hilbert 4`
Using `6` takes  2+ minutes to fully draw. It is fun to try once.
Using `7` causes the code to mis-calculate, so the line drawing goes off the left side of the canvas.
And it takes forever to do it.

## Automate

A message box is display once the document has been created asking if you want to close the document.

An extra parameter can be passed in:

An integer value the determines the levels to draw [Hilbert curve]. The default value is `4`.

**Example:**

```sh
python -m start 4
```

### Dev Container

From current example folder.

```sh
python -m start 4
```

### Cross Platform

From current example folder.

```sh
python -m start 4
```

### Linux/Mac

```sh
python ./ex/auto/draw/odev_draw_hilbert/start.py 4
```

### Windows

```ps
python .\ex\auto\draw\odev_draw_hilbert\start.py 4
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_hilbert'
```

This will copy the `odev_draw_hilbert` example to the examples folder.

In the terminal run:

```bash
cd odev_draw_hilbert
python -m start
```

[Hilbert curve]: https://en.wikipedia.org/wiki/Hilbert_curve
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/