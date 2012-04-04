# DocToPlainText

a simple and not so fail proof c# console application, that is able to convert a winword document into a plain text file. 
It is meant to be used in a script that parses a word document.

## Dependencies
_Microsoft.Office.Interop.Word_

The solution depends on an installed word application (it has been tested with winword 2007). The program will start the winword.exe, load the file and invoke the save as command.

## Usage
_DocToPlainText.exe c:/input.doc c:/output.txt_

_DocToPlainText.exe c:/input.doc_

When omitting the second parameter the application will write the contents of the wordfile to the console. This will only work with source files using ASCII characterset.