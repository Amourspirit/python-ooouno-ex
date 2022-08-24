# Speak Text

Demonstrates how to control Writer application's visible cursor.
Moves cursor by paragraph and sentence while reading out loud each sentence using [text-to-speech].

See [Extract Writer Text] for an example how to silence ODEV extra terminal output.

## Requirements

This example has additional [requirements](./requirements.txt) to install.

```text
pip install -r requirements.txt
```

see [requirements.txt](./requirements.txt)

## See

See Also:

- [Text API Overview]
- [Extract Writer Text]
- [using and comparing cursors]
- [OOO Development Tools]

See [source code](./start.py)

For advanced Text-to-Speech generation see [coqui-ai TTS](https://github.com/coqui-ai/TTS)

## Automate

### Cross Platform

From project root folder.

```shell
python -m main auto -p "ex/auto/writer/odev_speak/start.py --file resources/odt/cicero_dummy.odt"
```

### Linux

Run from current example folder.

```shell
python start.py --file "../../../../resources/odt/cicero_dummy.odt"
```

### Output

```text
Loading Office...
Opening /home/paul/Documents/Projects/Python/LibreOffice/ooouno_ex/resources/odt/cicero_dummy.odt
P<Cicero>
S<Cicero>
Sentence cursor passed end of current paragraph
P<Dummy Text>
S<Dummy Text>
Sentence cursor passed end of current paragraph
P<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful.>
S<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the masterbuilder of human happiness>
S<No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful>
Sentence cursor passed end of current paragraph
P<>
P<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure?>
S<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure>
S<To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it>
S<But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure>
Sentence cursor passed end of current paragraph
P<>
P<On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish.>
S<On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain>
S<These cases are perfectly simple and easy to distinguish>
Sentence cursor passed end of current paragraph
P<>
P<In a free hour, when our power of choice is untrammelled and when nothing prevents our being able to do what we like best, every pleasure is to be welcomed and every pain avoided. But in certain circumstances and owing to the claims of duty or the obligations of business it will frequently occur that pleasures have to be repudiated and annoyances accepted. The wise man therefore always holds in these matters to this principle of selection: he rejects pleasures to secure other greater pleasures, or else he endures pains to avoid worse pains.>
S<In a free hour, when our power of choice is untrammelled and when nothing prevents our being able to do what we like best, every pleasure is to be welcomed and every pain avoided>
S<But in certain circumstances and owing to the claims of duty or the obligations of business it will frequently occur that pleasures have to be repudiated and annoyances accepted>
S<The wise man therefore always holds in these matters to this principle of selection he rejects pleasures to secure other greater pleasures, or else he endures pains to avoid worse pains>
Sentence cursor passed end of current paragraph
P<>
P<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful.>
S<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the masterbuilder of human happiness>
S<No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful>
Sentence cursor passed end of current paragraph
P<>
P<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure? On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish. In a free hour, when our power of choice is untrammelled and when nothing prevents our>
S<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure>
S<To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it>
S<But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure>
S<On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain>
S<These cases are perfectly simple and easy to distinguish>
S<In a free hour, when our power of choice is untrammelled and when nothing prevents our>
Sentence cursor passed end of current paragraph
Closing the document
Closing Office
Office terminated
Office bridge has gone!!
```

[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html
[using and comparing cursors]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html#using-and-comparing-text-cursors
[Extract Writer Text]: ../odev_doc_print_console/
[text-to-speech]: https://pypi.org/project/text-to-speech/
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
