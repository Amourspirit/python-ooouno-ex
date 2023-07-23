# Extract Graphics

<p align="center">
  <img width="200" src="https://user-images.githubusercontent.com/4193389/202314276-77dfb4ac-0a44-451b-a9b6-f8c758198e4b.svg">
</p>

This example use the output file from [odev_build_doc](../odev_build_doc/) and extracts graphics into specified directory (or tmp directory if not specified).

This demo uses [OOO Development Tools]

## See

See Also:

- [OOO Development Tools - Chapter 8. Graphic Content]
- [8.3 Accessing Linked Images and Shapes]

## Automate

### Cross Platform

From this folder.

```sh
python -m start --file "../../../../resources/odt/build.odt"
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/writer/odev_extract_graphics/start.py --file "resources/odt/build.odt"
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_extract_graphics\start.py --file "resources/odt/build.odt"
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_extract_graphics'
```

This will copy the `odev_extract_graphics` example to the examples folder.

In the terminal run:

```bash
cd odev_extract_graphics
python -m start
```


## Output

```text
No. of text graphics: 2
Saving graphic in 'C:\Users\user\AppData\Local\Temp\tmpixludwxs\graphics0.png'
Image size in pixels: 319 X 274
Saving graphic in 'C:\Users\user\AppData\Local\Temp\tmpixludwxs\graphics1.png'
Image size in pixels: 319 X 274

Could not obtain text shapes supplier

No. of draw shapes: 5
Shape Name: Shape1
  Type: com.sun.star.drawing.GraphicObjectShape
  Point (mm): [0, 0]
  Size (mm): [61, 58]
Shape Name: Shape2
  Type: com.sun.star.drawing.LineShape
  Point (mm): [0, 0]
  Size (mm): [88, 0]
Shapes does not have a name property
  Type: FrameShape
  Size (mm): [40, 0]
Shapes does not have a name property
  Type: FrameShape
  Size (mm): [61, 58]
Shapes does not have a name property
  Type: FrameShape
  Size (mm): [91, 86]
```

[8.3 Accessing Linked Images and Shapes]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter08.html#accessing-linked-images-and-shapes
[OOO Development Tools - Chapter 8. Graphic Content]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter08.html
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
