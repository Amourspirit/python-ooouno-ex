# Styles Info

Report all the style family names used in a document.
For a text document there are five style family names.
List the container names used by each style family.

Print text document style family + "Standard" container property info on:

- page styles
- paragraph styles    -- lots of container names (e.g. "Header 1", "Numbering 1 Start")

## See

See Also:

- [Text Styles]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Dev Container

From this folder.

```sh
python -m start --file "./data/cicero_dummy.odt"
```

### Cross Platform

From this folder.

```sh
python -m start --file "./data/cicero_dummy.odt"
```

### Linux/Mac

From project root folder (for default document).

```sh
python ./ex/auto/writer/odev_styles_info/start.py
```

### Windows

From project root folder (for default document).

```ps
python .\ex\auto\writer\odev_styles_info\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_styles_info'
```

This will copy the `odev_styles_info` example to the examples folder.

In the terminal run:

```bash
cd odev_styles_info
python -m start --show --verbose  --file "./data/cicero_dummy.odt"
```

## Output

```text
No. of Style Family Names: 7
  CellStyles
  CharacterStyles
  FrameStyles
  NumberingStyles
  PageStyles
  ParagraphStyles
  TableStyles

0 "CellStyles" Style Family contains containers:
  No names found
1 "CharacterStyles" Style Family contains containers:
No. of names: 27
  ---------------------------|----------------------------|----------------------------|----------------------------
  Bullet Symbols             | Caption characters         | Citation                   | Definition
  Drop Caps                  | Emphasis                   | Endnote anchor             | Endnote Symbol   
  Example                    | Footnote anchor            | Footnote Symbol            | Index Link       
  Internet link              | Line numbering             | Main index entry           | Numbering Symbols
  Page Number                | Placeholder                | Rubies                     | Source Text      
  Standard                   | Strong Emphasis            | Teletype                   | User Entry       
  Variable                   | Vertical Numbering Symbols | Visited Internet Link      |



2 "FrameStyles" Style Family contains containers:
No. of names: 7
  -----------|------------|------------|------------
  Formula    | Frame      | Graphics   | Labels
  Marginalia | OLE        | Watermark  |



3 "NumberingStyles" Style Family contains containers:
No. of names: 11
  --------------|---------------|---------------|---------------
  List 1        | List 2        | List 3        | List 4
  List 5        | No List       | Numbering 123 | Numbering ABC
  Numbering abc | Numbering IVX | Numbering ivx |



4 "PageStyles" Style Family contains containers:
No. of names: 10
  -----------|------------|------------|------------
  Endnote    | Envelope   | First Page | Footnote
  HTML       | Index      | Landscape  | Left Page
  Right Page | Standard   |            |



5 "ParagraphStyles" Style Family contains containers:
No. of names: 125
  ---------------------|----------------------|----------------------|----------------------
  Addressee            | Appendix             | Bibliography 1       | Bibliography Heading
  Caption              | Contents 1           | Contents 10          | Contents 2
  Contents 3           | Contents 4           | Contents 5           | Contents 6
  Contents 7           | Contents 8           | Contents 9           | Contents Heading
  Drawing              | Endnote              | Figure               | Figure Index 1
  Figure Index Heading | First line indent    | Footer               | Footer left
  Footer right         | Footnote             | Frame contents       | Hanging indent
  Header               | Header and Footer    | Header left          | Header right
  Heading              | Heading 1            | Heading 10           | Heading 2
  Heading 3            | Heading 4            | Heading 5            | Heading 6
  Heading 7            | Heading 8            | Heading 9            | Horizontal Line
  Illustration         | Index                | Index 1              | Index 2
  Index 3              | Index Heading        | Index Separator      | List
  List 1               | List 1 Cont.         | List 1 End           | List 1 Start
  List 2               | List 2 Cont.         | List 2 End           | List 2 Start
  List 3               | List 3 Cont.         | List 3 End           | List 3 Start
  List 4               | List 4 Cont.         | List 4 End           | List 4 Start
  List 5               | List 5 Cont.         | List 5 End           | List 5 Start
  List Contents        | List Heading         | List Indent          | Marginalia
  Numbering 1          | Numbering 1 Cont.    | Numbering 1 End      | Numbering 1 Start
  Numbering 2          | Numbering 2 Cont.    | Numbering 2 End      | Numbering 2 Start
  Numbering 3          | Numbering 3 Cont.    | Numbering 3 End      | Numbering 3 Start
  Numbering 4          | Numbering 4 Cont.    | Numbering 4 End      | Numbering 4 Start
  Numbering 5          | Numbering 5 Cont.    | Numbering 5 End      | Numbering 5 Start
  Object index 1       | Object index heading | Preformatted Text    | Quotations
  Salutation           | Sender               | Signature            | Standard
  Subtitle             | Table                | Table Contents       | Table Heading
  Table index 1        | Table index heading  | Text                 | Text body
  Text body indent     | Title                | User Index 1         | User Index 10
  User Index 2         | User Index 3         | User Index 4         | User Index 5
  User Index 6         | User Index 7         | User Index 8         | User Index 9
  User Index Heading   |                      |                      |



6 "TableStyles" Style Family contains containers:
No. of names: 11
  --------------------|---------------------|---------------------|---------------------
  Academic            | Box List Blue       | Box List Green      | Box List Red
  Box List Yellow     | Default Style       | Elegant             | Financial
  Simple Grid Columns | Simple Grid Rows    | Simple List Shaded  |



ParagraphStyles "Standard" Properties
  BorderDistance: 0
  BottomBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  BottomBorderDistance: 0
  BreakType: <Enum instance com.sun.star.style.BreakType ('NONE')>
  Category: 4
  CharAutoKerning: True
  CharBackColor: -1
  CharBackTransparent: True
  CharBorderDistance: 0
  CharBottomBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  CharBottomBorderDistance: 0
  CharCaseMap: 0
  CharColor: -1
  CharCombineIsOn: False
  CharCombinePrefix:
  CharCombineSuffix:
  CharContoured: False
  CharCrossedOut: False
  CharDiffHeight: 0.0
  CharDiffHeightAsian: 0.0
  CharDiffHeightComplex: 0.0
  CharEmphasis: 0
  CharEscapement: 0
  CharEscapementHeight: 100
  CharFlash: False
  CharFontCharSet: 1
  CharFontCharSetAsian: 1
  CharFontCharSetComplex: 1
  CharFontFamily: 3
  CharFontFamilyAsian: 6
  CharFontFamilyComplex: 6
  CharFontName: Liberation Serif
  CharFontNameAsian: Noto Serif CJK SC
  CharFontNameComplex: Lohit Devanagari
  CharFontPitch: 2
  CharFontPitchAsian: 2
  CharFontPitchComplex: 2
  CharFontStyleName: 
  CharFontStyleNameAsian:
  CharFontStyleNameComplex:
  CharHeight: 12.0
  CharHeightAsian: 10.5
  CharHeightComplex: 12.0
  CharHidden: False
  CharHighlight: -1
  CharInteropGrabBag: ()
  CharKerning: 0
  CharLeftBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  CharLeftBorderDistance: 0
  CharLocale: (com.sun.star.lang.Locale){ Language = (string)"en", Country = (string)"CA", Variant = (string)"" }
  CharLocaleAsian: (com.sun.star.lang.Locale){ Language = (string)"zh", Country = (string)"CN", Variant = (string)"" }
  CharLocaleComplex: (com.sun.star.lang.Locale){ Language = (string)"hi", Country = (string)"IN", Variant = (string)"" }
  CharNoHyphenation: True
  CharOverline: 0
  CharOverlineColor: -1
  CharOverlineHasColor: False
  CharPosture: <Enum instance com.sun.star.awt.FontSlant ('NONE')>
  CharPostureAsian: <Enum instance com.sun.star.awt.FontSlant ('NONE')>
  CharPostureComplex: <Enum instance com.sun.star.awt.FontSlant ('NONE')>
  CharPropHeight: 100
  CharPropHeightAsian: 100
  CharPropHeightComplex: 100
  CharRelief: 0
  CharRightBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  CharRightBorderDistance: 0
  CharRotation: 0
  CharRotationIsFitToLine: False
  CharScaleWidth: 100
  CharShadingValue: 0
  CharShadowFormat: (com.sun.star.table.ShadowFormat){ Location = (com.sun.star.table.ShadowLocation)NONE, ShadowWidth = (short)0xb0, IsTransparent = (boolean)false, Color = (long)0x808080 }
  CharShadowed: False
  CharStrikeout: 0
  CharTopBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  CharTopBorderDistance: 0
  CharTransparence: 100
  CharUnderline: 0
  CharUnderlineColor: -1
  CharUnderlineHasColor: False
  CharWeight: 100.0
  CharWeightAsian: 100.0
  CharWeightComplex: 100.0
  CharWordMode: False
  DisplayName: Header
  DropCapCharStyleName:
  DropCapFormat: (com.sun.star.style.DropCapFormat){ Lines = (byte)0x0, Count = (byte)0x0, Distance = (short)0x0 }
  DropCapWholeWord: False
  FillBackground: False
  FillBitmap: None
  FillBitmapLogicalSize: True
  FillBitmapMode: <Enum instance com.sun.star.drawing.BitmapMode ('REPEAT')>
  FillBitmapName: 
  FillBitmapOffsetX: 0
  FillBitmapOffsetY: 0
  FillBitmapPositionOffsetX: 0
  FillBitmapPositionOffsetY: 0
  FillBitmapRectanglePoint: <Enum instance com.sun.star.drawing.RectanglePoint ('MIDDLE_MIDDLE')>
  FillBitmapSizeX: 0
  FillBitmapSizeY: 0
  FillBitmapStretch: True
  FillBitmapTile: True
  FillBitmapURL: None
  FillColor: 7512015
  FillColor2: 7512015
  FillGradient: (com.sun.star.awt.Gradient){ Style = (com.sun.star.awt.GradientStyle)LINEAR, StartColor = (long)0x0, EndColor = (long)0xffffff, Angle = (short)0x0, Border = (short)0x0, XOffset = (short)0x32, YOffset = (short)0x32, StartIntensity = (short)0x64, EndIntensity = (short)0x64, StepCount = (short)0x0 }
  FillGradientName: 
  FillGradientStepCount: 0
  FillHatch: (com.sun.star.drawing.Hatch){ Style = (com.sun.star.drawing.HatchStyle)SINGLE, Color = (long)0x3465a4, Distance = (long)0x14, Angle = (long)0x0 }
  FillHatchName:
  FillStyle: <Enum instance com.sun.star.drawing.FillStyle ('NONE')>
  FillTransparence: 0
  FillTransparenceGradient: (com.sun.star.awt.Gradient){ Style = (com.sun.star.awt.GradientStyle)LINEAR, StartColor = (long)0x0, EndColor = (long)0x0, Angle = (short)0x0, Border = (short)0x0, XOffset = (short)0x32, YOffset = (short)0x32, StartIntensity = (short)0x64, EndIntensity = (short)0x64, StepCount = (short)0x0 }
  FillTransparenceGradientName:
  FollowStyle: Header
  Hidden: False
  IsAutoUpdate: False
  IsPhysical: True
  LeftBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  LeftBorderDistance: 0
  LinkStyle:
  NumberingLevel: 0
  NumberingStyleName: 
  OutlineLevel: 0
  PageDescName: None
  PageNumberOffset: None
  ParaAdjust: 0
  ParaBackColor: -1
  ParaBackGraphic: None
  ParaBackGraphicFilter:
  ParaBackGraphicLocation: <Enum instance com.sun.star.style.GraphicLocation ('NONE')>
  ParaBackGraphicURL: None
  ParaBackTransparent: True
  ParaBottomMargin: 0
  ParaBottomMarginRelative: 100
  ParaContextMargin: False
  ParaExpandSingleWord: False
  ParaFirstLineIndent: 0
  ParaFirstLineIndentRelative: 100
  ParaHyphenationMaxHyphens: 0
  ParaHyphenationMaxLeadingChars: 2
  ParaHyphenationMaxTrailingChars: 2
  ParaHyphenationNoCaps: False
  ParaInteropGrabBag: ()
  ParaIsAutoFirstLineIndent: False
  ParaIsCharacterDistance: True
  ParaIsConnectBorder: True
  ParaIsForbiddenRules: True
  ParaIsHangingPunctuation: True
  ParaIsHyphenation: False
  ParaKeepTogether: False
  ParaLastLineAdjust: 0
  ParaLeftMargin: 0
  ParaLeftMarginRelative: 100
  ParaLineNumberCount: False
  ParaLineNumberStartValue: 0
  ParaLineSpacing: (com.sun.star.style.LineSpacing){ Mode = (short)0x0, Height = (short)0x64 }
  ParaOrphans: 2
  ParaRegisterModeActive: False
  ParaRightMargin: 0
  ParaRightMarginRelative: 100
  ParaShadowFormat: (com.sun.star.table.ShadowFormat){ Location = (com.sun.star.table.ShadowLocation)NONE, ShadowWidth = (short)0xb0, IsTransparent = (boolean)false, Color = (long)0x808080 }
  ParaSplit: True
  ParaTabStops: ((com.sun.star.style.TabStop){ Position = (long)0x225b, Alignment = (com.sun.star.style.TabAlign)CENTER, DecimalChar = (char)'.', FillChar = (char)' ' }, (com.sun.star.style.TabStop){ Position = (long)0x44b6, Alignment = (com.sun.star.style.TabAlign)RIGHT, DecimalChar = (char)'.', FillChar = (char)' ' })
  ParaTopMargin: 0
  ParaTopMarginRelative: 100
  ParaUserDefinedAttributes: pyuno object (com.sun.star.container.XNameContainer)0x2599f4aa498{implementationName=SvUnoAttributeContainer, supportedServices={com.sun.star.xml.AttributeContainer}, supportedInterfaces={com.sun.star.lang.XServiceInfo,com.sun.star.lang.XUnoTunnel,com.sun.star.container.XNameContainer,com.sun.star.lang.XTypeProvider,com.sun.star.uno.XWeak,com.sun.star.uno.XAggregation}}
  ParaVertAlignment: 0
  ParaWidows: 2
  RightBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  RightBorderDistance: 0
  Rsid: 0
  SnapToGrid: True
  StyleInteropGrabBag: ()
  TopBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
  TopBorderDistance: 0
  WritingMode: 4
```

[Text Styles]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter06.html
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
