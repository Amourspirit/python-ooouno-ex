# Show Book

Another approach for moving around a document involves the `XEnumerationAccess` interface which treats the document as a series of Paragraph text contents.

This example prints all the text in every paragraph using [OOO Development Tools] and enumeration access.

## See

See Also:

- [Text API Overview]
- [Treating a Document as Paragraphs and Text Portions]
- [Class LoEvents]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Cross Platform

From this folder.

```sh
python -m start --file "./data/cicero_dummy.odt"
```

### Linux/Mac

Run from current example folder.

From project root folder (default file).

```sh
python ./ex/auto/writer/odev_show_book/start.py
```

### Windows

From project root folder (default file).

```ps
python .\ex\auto\writer\odev_show_book\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_show_book'
```

This will copy the `odev_show_book` example to the examples folder.

In the terminal run:

```bash
cd odev_show_book
python -m start
```

## Output

```text
P--
  Text = "Cicero"
P--
  Text = "Dummy Text"
P--
  Text = "But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful."
P--
  Text = ""
P--
  Text = "Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure?"
P--
  Text = ""
P--
  Text = "On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish."
P--
  Text = ""
P--
  Text = "In a free hour, when our power of choice is untrammelled and when nothing prevents our being able to do what we like best, every pleasure is to be welcomed and every pain avoided. But in certain circumstances and owing to the claims of duty or the obligations of business it will frequently occur that pleasures have to be repudiated and annoyances accepted. The wise man therefore always holds in these matters to this principle of selection: he rejects pleasures to secure other greater pleasures, or else he endures pains to avoid worse pains."
P--
  Text = ""
P--
  Text = "But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful."
P--
  Text = ""
P--
  Text = "Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure? On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish. In a free hour, when our power of choice is untrammelled and when nothing prevents our"
```

[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html

[Treating a Document as Paragraphs and Text Portions]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html#treating-a-document-as-paragraphs-and-text-portions

[Class LoEvents]: http://localhost:8000/docs/_build/html/src/events/lo_events/lo_events.html

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
