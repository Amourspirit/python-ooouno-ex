# Impress append Slides to existing slide show

<p align="center">
  <img width="435" height="448" src="https://user-images.githubusercontent.com/4193389/198401485-94062f29-6a24-40f7-8873-fce8abaff481.png">
</p>

This example demonstrates how to combine Slide show documents using Impress.

There is one limitation at this time.
For every slide that is appended the user is forced to click a **yes** ( five time when adding `points.odp` ) in a dialog prompt.

There is a remedy for this but it is outside of the scope for this demo, mainly
due to the many variations it will be up to end user to make a custom implementation

One potential solution would be [autopy](https://pypi.org/project/autopy/)
however `autopy` is for `X11` on Linux and will not work for `Wayland`.
There are other Wayland solutions available.
At this time there does not seem to be a solution that works for both X11 and Wayland.

This demo uses [OOO Development Tools]

## Automate

An extra parameters can be passed in:

The first parameter would be the slide show file to append to.

All successive files are append to the first.

**Example:**

```sh
python ./ex/auto/impress/odev_append_slides/start.py "resources/presentation/algs.odp" "resources/presentation/points.odp"
```

If no args are passed in then the `points.odp` is appended to `algs.odp`.

The document is not saved by default.

A message box is display once the document has been created asking if you want to close the document.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_append_slides/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_append_slides\start.py
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/