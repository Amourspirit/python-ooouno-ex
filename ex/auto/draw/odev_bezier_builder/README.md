# Draw Bezier Curve

![Screen_shot](https://user-images.githubusercontent.com/4193389/198354784-e14025b0-67a5-4d63-b95a-543a414384fa.png)

Demonstrates reading a text file contains Bezier curve data that is recreated in a Draw document.

A message box is display once the document has been created asking if you want to close the document.

This demo uses [OOO Development Tools]

See [OOO Development Tools: Chapter 15. Complex Shapes](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/chapter15.html)

## Automate

An extra parameter can be passed in:

A value between `0` and `3` The default value is `2`.
Each value represents a different Bezier curve file.

**Example:**

```sh
python -m start 1
```

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/draw/odev_bezier_builder/start.py
```

### Windows

```ps
python .\ex\auto\draw\odev_bezier_builder\start.py
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/