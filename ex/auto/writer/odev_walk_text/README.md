<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/181919190-a415dc2d-762c-48a6-b660-cbfc92b46db1.gif" alt="animation"/>
</p>

# Walk Text

Using [OOO Development Tools] (ODEV) this example demonstrates how to control Writer application's visible cursor.
Moves cursor by paragraph, line and word.

See [Extract Writer Text] for an example how to silence ODEV extra terminal output.

## See

See Also:

- [Text API Overview]
- [Extract Writer Text]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Cross Platform

From project root folder.

```shell
python -m main auto -p "ex/auto/writer/odev_walk_text/start.py --file resources/odt/cicero_dummy.odt"
```

### Linux

Run from current example folder.

```shell
python start.py --file "../../../../resources/odt/cicero_dummy.odt"
```

### Output
```text
P<Cicero>
P<Dummy Text>
P<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful.>
P<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure?>
P<On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish.>
P<In a free hour, when our power of choice is untrammelled and when nothing prevents our being able to do what we like best, every pleasure is to be welcomed and every pain avoided. But in certain circumstances and owing to the claims of duty or the obligations of business it will frequently occur that pleasures have to be repudiated and annoyances accepted. The wise man therefore always holds in these matters to this principle of selection: he rejects pleasures to secure other greater pleasures, or else he endures pains to avoid worse pains.>
P<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful.>
P<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure? On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. These cases are perfectly simple and easy to distinguish. In a free hour, when our power of choice is untrammelled and when nothing prevents our>
Word count: 603
L<Cicero>
L<Dummy Text>
L<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born >
L<and I will give you a complete account of the system, and expound the actual teachings of the great >
L<explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids >
L<pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure >
L<rationally encounter consequences that are extremely painful.>
L<>
L<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but >
L<because occasionally circumstances occur in which toil and pain can procure him some great pleasure. >
L<To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain >
L<some advantage from it? But who has any right to find fault with a man who chooses to enjoy a >
L<pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant >
L<pleasure?>
L<>
L<On the other hand, we denounce with righteous indignation and dislike men who are so beguiled and >
L<demoralized by the charms of pleasure of the moment, so blinded by desire, that they cannot foresee >
L<the pain and trouble that are bound to ensue; and equal blame belongs to those who fail in their duty >
L<through weakness of will, which is the same as saying through shrinking from toil and pain. These >
L<cases are perfectly simple and easy to distinguish.>
L<>
L<In a free hour, when our power of choice is untrammelled and when nothing prevents our being able to >
L<do what we like best, every pleasure is to be welcomed and every pain avoided. But in certain >
L<circumstances and owing to the claims of duty or the obligations of business it will frequently occur >
L<that pleasures have to be repudiated and annoyances accepted. The wise man therefore always holds in >
L<these matters to this principle of selection: he rejects pleasures to secure other greater pleasures, or else >
L<he endures pains to avoid worse pains.>
L<>
L<But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born >
L<and I will give you a complete account of the system, and expound the actual teachings of the great >
L<explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids >
L<pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure >
L<rationally encounter consequences that are extremely painful.>
L<>
L<Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but >
L<because occasionally circumstances occur in which toil and pain can procure him some great pleasure. >
L<To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain >
L<some advantage from it? But who has any right to find fault with a man who chooses to enjoy a >
L<pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant >
L<pleasure? On the other hand, we denounce with righteous indignation and dislike men who are so >
L<beguiled and demoralized by the charms of pleasure of the moment, so blinded by desire, that they >
L<cannot foresee the pain and trouble that are bound to ensue; and equal blame belongs to those who fail >
L<in their duty through weakness of will, which is the same as saying through shrinking from toil and >
L<pain. These cases are perfectly simple and easy to distinguish. In a free hour, when our power of choice >
L<is untrammelled and when nothing prevents our>
Closing the document
Closing Office
Office terminated
Office bridge has gone!!
```

[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html
[Extract Writer Text]: ../odev_doc_print_console/
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
