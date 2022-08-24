<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/180623026-9e5b96fc-22c0-43b8-a612-139eb3b28737.png" alt="dispatch"/>
</p>

# Dispatch Commands Example

This example demonstrates how to dispatch commands using [OOO Development Tools] (ODEV).

LibreOffice has a comprehensive web page listing all the dispatch commands [Development/DispatchCommands](https://wiki.documentfoundation.org/Development/DispatchCommands).

This example also demonstrates hooking ODEV events that in this case allow for finer control over which commands are dispatched.

See Also: [Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)

See [source code](./start.py)

## Automate

Running the following command opens a Write document, puts it into read-only mode.
Next a Get Involved wep page is opened in local web browser.

### Cross Platform

From project root folder.

```shell
python -m main auto -p "ex/auto/general/odev_dispatch/start.py -d resources/odt/story.odt"
```

### Linux

Run from current example folder.

```shell
python start.py -d "../../../../resources/odt/story.odt"
```

### Example console output

```text
PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process 'ex\auto\general\odev_dispatch\start.py -d "resources\odt\story.odt"'
Loading Office...
Opening D:\Users\user\Python\python-ooouno-ex\resources\odt\story.odt
Dispatching: ReadOnlyDoc
Dispatched: ReadOnlyDoc
Dispatching: GetInvolved
Dispatched: GetInvolved
About dispatch canceled
```

[OOO Development Tools]:https://python-ooo-dev-tools.readthedocs.io/en/latest/
