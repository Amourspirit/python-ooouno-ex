# Impress Extract Text

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198415603-a7ea1593-06a7-482f-b245-0933d0f5950d.png" width="396" height="314">
</p>

Attempts to extract the text from the slide deck.

The order of the text extracted may not be the same as the order
that it appears in the file; it depends on the order that the text shapes are saved inside the file.

Connection to LibreOffice is set to `headless=True` (invisible).
Do not need to display LibreOffice when just writing to console.
In some cases headless mode is required such as when running on a Headless server.

This demo uses [OOO Development Tools]

See Also:

- [OOO Development Tools - Chapter 17. Slide Deck Manipulation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/chapter17.html)

## Automate

A single parameters can be passed in which is the slide show document to read from:

**Example:**

```sh
python ./ex/auto/impress/odev_extract_text/start.py "resources/presentation/algs.odp"
```

If no parameters are passed then the script is run with the above parameters.

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
python ./ex/auto/impress/odev_extract_text/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_extract_text\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_extract_text'
```

This will copy the `odev_extract_text` example to the examples folder.

In the terminal run:

```bash
cd odev_extract_text
python -m start
```

## OUTPUT

This is the output from `tests/fixtures/presentation/algs.odp`

```text
Loading Office...
Opening /home/user/Projects/ooouno-dev-tools/tests/fixtures/presentation/algs.odp
-----------------Text Content-----------------
An algorithm is a finite set of unambiguous instructions for solving a problem.
An algorithm is correct if on all legitimate inputs, it outputs the right answer in a finite amount of time

Can be expressed as
pseudocode
flow charts
text in a natural language (e.g. English)
computer code
What is a Algorithm?
The theoretical  study of how to solve  computational problems
sorting a list of numbers
finding a shortest route on a map
scheduling when to work on homework
answering web search queries
and so on...

Algorithm Design

Their impact is broad and far-reaching.
Internet. Web search, packet routing, distributed file sharing, ...
Biology. Human genome project, protein folding, ...
Computers. Circuit layout, file system, compilers, ...
Computer graphics. Movies, video games, virtual reality, ...
Security. Cell phones, e-commerce, voting machines, ...
Multimedia. MP3, JPG, DivX, HDTV, face recognition, ...
Social networks.  Recommendations, news feeds, advertisements, ...
Physics. N-body simulation, particle collision simulation, ...
The Importance of Algorithms
Ten algorithms having "the greatest influence on the development and practice of science and engineering in the 20th century".
Dongarra and Sullivan
Top Ten Algorithms of the Century
Computing in Science and Engineering
January/February 2000

Barry Cipra
The Best of the 20th Century: Editors Name Top 10 Algorithms
SIAM News
Volume 33, Number 4, May 2000
<http://www.siam.org/pdf/news/637.pdf>
The Top 10 Algorithms of the 20th Century
1946: The Metropolis (Monte Carlo) Algorithm.
Uses random processes to find answers to problems that are too complicated to solve exactly.

1947: Simplex Method for Linear Programming.
A fast technique for maximizing or minimizing a linear function of several variables, applicable to planning and decision-making.

1950: Krylov Subspace Iteration Method.
A technique for rapidly solving the linear equations that are common in scientific computation.
What are the Top 10?
1951: The Decompositional Approach to Matrix Computations. A collection of techniques for numerical linear algebra.

1957: The Fortran Optimizing Compiler.

1959: QR Algorithm for Computing Eigenvalues.
A crucial matrix operation made swift and practical. Application areas include computer vision, vibration analysis, data analysis.

1962: Quicksort Algorithm.
We will look at this.
1965: Fast Fourier Transform (FFT). It breaks down waveforms (like sound) into periodic components. Used in many different areas (e.g. digital signal processing , solving partial differential equations, fast multiplication of large integers.)

1977: Integer Relation Detection. A fast method for finding simple equations that explain collections of data.

1987: Fast Multipole Method. Deals with the complexity of n-body calculations. It is applied in problems ranging from celestial mechanics to protein folding.
Introduction to Algorithms
Thomas Cormen, Charles Leiserson, Ronald Rivest, Clifford Stein
McGraw Hill, 2003, 2nd edition
mathematical, advanced, the standard text
now up to version 3 (MIT)

lots of resources online;
see video section

Books
continued

Algorithms
Robert Sedgewick, Kevin Wayne
Addison-Wesley, 2011, 4th ed.
implementation (Java) and theory
intermediate level

Data Structures and Algorithms in Java
Robert Lafore
Sams Publishing, 2002, 2nd ed.
Java examples; old
basic level; not much analysis

Algorithms Unlocked
Thomas H. Cormen
MIT Press, March 2013

Nine Algorithms that Changed the Future
John MacCormick
Princeton University Press, 2011
<http://users.dickinson.edu/~jmac/9algorithms/>
search engine indexing, pagerank, public key cryptography, error-correcting codes, pattern recognition, data compression, databases, digital signatures, computablity
Fun Overviews

Algorithmic Puzzles
Anany Levitin, Maria Levitin
Oxford University Press, , 2011

Algorithmics: The Spirit of Computing
David Harel, Yishai Feldman
Addison-Wesley; 3 ed., 2004
(and Springer, 2012)
<http://www.wisdom.weizmann.ac.il/~harel/algorithmics.html>

The New Turing Omnibus: Sixty-Six
Excursions in Computer Science
A. K. Dewdney
Holt, 1993
66 short article; e.g. detecting primes, noncomputable functions, self-replicating computers, fractals, genetic algorithms, Newton-Raphson Method, viruses

----------------------------------------------
Closing the document
Closing Office
Office terminated
Office bridge has gone!!
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/