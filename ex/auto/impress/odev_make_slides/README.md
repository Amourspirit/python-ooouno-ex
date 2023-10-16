# Impress Make Slides

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/200715432-e781ac26-beed-48c9-8a01-a9a4b58d7720.png" width="640" height="285">
</p>

Demonstrates Creating making slides.

Example creates a deck of four or five slides, illustrating different aspects of slide generation:

- Slide 1. A slide combining a title and subtitle,
- Slide 2. A slide with a title, bullet points, and an image.
- Slide 3. A slide with a title, and an embedded video which plays automatically when that slide appears during a slide show.
- Slide 4. A slide with an ellipse and a rounded rectangle acting as buttons. During a slide show, clicking on the ellipse starts a video playing in an external viewer. Clicking on the rounded rectangle causes the slide show to jump to the first slide in the deck.
- Slide 5. (Windows only) This slide contains eight shapes generated using dispatches, including special symbols, block arrows, 3D shapes, flowchart elements, call-outs, and stars.

This demo uses [OOO Development Tools]

Creation of slide five requires GUI automation provided by [ooo-dev-tools-gui-win], also available as a [LibreOffice Extension](https://extensions.libreoffice.org/en/extensions/show/41986) for Windows.

See Also:

- [OOO Development Tools - Part 3: Draw & Impress](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/index.html)

## Automate

A message box is display once the slide show has ended asking if you want to close the document.

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
python ./ex/auto/impress/odev_make_slides/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_make_slides\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_make_slides'
```

This will copy the `odev_make_slides` example to the examples folder.

In the terminal run:

```bash
cd odev_make_slides
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[ooo-dev-tools-gui-win]: https://ooo-dev-tools-gui-win.readthedocs.io/en/latest/index.html
