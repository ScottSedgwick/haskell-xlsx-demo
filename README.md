# xlsxDemo

[![Hackage](https://img.shields.io/hackage/v/xlsxDemo.svg?logo=haskell)](https://hackage.haskell.org/package/xlsxDemo)
[![MIT license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

An example of how to use the Haskell Xlsx library.

This example:

* Reads a template file (called `template1.xlsx`)
* Reads a second template file (called `template2.xlsx`)
* Copies the value from template1, Sheet1, Cell B3 to template2, Sheet1, Cell C4
* Writes the updated spreadsheet to an output file (called `output.xlsx`)

It is not elegant or clean.  Work is required.

All the code is in `app/Main.hs`.

Build the code using `cabal build`.

Run the code using `cabal exec xlsxDemo`.
