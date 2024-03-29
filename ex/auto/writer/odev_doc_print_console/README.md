# Extract Writer Text

This example shows how to print the contents of a Writer document to the console/terminal.

This example uses [OOO Development Tools] (OooDev).

## Events

This example also use OooDev's events to hook and suppress internal printing via [Class LoEvents].

## See

See Also:

- [Text API Overview]
- [Class LoEvents]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Dev Container

From project root folder.

```sh
python -m start --file "./data/cicero_dummy.odt"
```

### Cross Platform

From project root folder.

```sh
python -m start --file "./data/cicero_dummy.odt"
```

### Linux/Mac

Run from current example folder.

```sh
python ./ex/auto/writer/odev_doc_print_console/start.py --file "ex/auto/writer/odev_doc_print_console/data/cicero_dummy.odt"
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_doc_print_console\start.py --file "ex/auto/writer/odev_doc_print_console/data/cicero_dummy.odt"
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_doc_print_console'
```

This will copy the `odev_doc_print_console` example to the examples folder.

In the terminal run:

```bash
cd odev_doc_print_console
python -m start
```

### Output

```text
>>> python -m main auto --process 'ex/auto/writer/odev_doc_print_console/start.py --file "resources/odt/cicero_dummy.odt"'
-------------------Text Content-------------------
Cicero
Dummy Text
But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful.

Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure?

On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish.

In a free hour, when our power of choice is untrammelled and when nothing prevents our being able to do what we like best, every pleasure is to be welcomed and every pain avoided. But in certain circumstances and owing to the claims of duty or the obligations of business it will frequently occur that pleasures have to be repudiated and annoyances accepted. The wise man therefore always holds in these matters to this principle of selection: he rejects pleasures to secure other greater pleasures, or else he endures pains to avoid worse pains.

But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful.

Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure? On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish. In a free hour, when our power of choice is untrammelled and when nothing prevents our
--------------------------------------------------
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html
[Class LoEvents]: http://localhost:8000/docs/_build/html/src/events/lo_events/lo_events.html

