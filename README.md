# underline-text-in-excel-cells

A VB script to automatically underline sub-strings in a range of Excel cells and a python script to create variations of name strings.

## VB script to underline sub-strings Excel cells

In Excel, it is not possible to underline (or change formatting) of a sub-string in a cell with search-and-replace, conditioanl formatting or formulae.
A workaround is to copy text to a word document do search-and-replace with formatting options, but the VB script in this repository provides a fully programatic one-click solution.

Sometimes, we need to underline specific sub-strings, like names, in an excel data sheet for administrative purposes. This VB script was conceived to automate this arduous task.

The VB script can be used not just for names, but for any other strings, and other formatting changes within cells can of course also be performed.

### Preparation

Adjust the sheet name and range in the VB script to match the cells that contain all the text to search for underlining.

```vb
Dim NamesVector
NamesVector = Create_Vector(Sheets("Sheet2").Range("a1:b4"))
```

In the example above, cells `a1:b4` in sheet `Sheet2` each contain text that should be underlined.

### Execution

Select a range of cells in the data where underlining should occur and run the script in the VBA editor of the Excel file.

## Python helper script to generate variations on name strings

This script is tailored specifically for administrative tasks in Japanese research institutions.

This script reads a simple two-column text file of member names and outputs a larger variety of name arrangements in csv format, which can be used with a VB script to underline these names in an Excel sheet.

The two-column text file should be named `member_names.txt` and be located in the same location as the python script. The two-column data is separated by a comma. Left column: kanji/kana name `<Surname> <First name>` (space between surname and first name). Right column: romaji name `<First names> <Surname>` (space between first name and surname). Middle names possible, but only for romaji name.

E.g.
```txt
佐藤 健二, Kenji Sato
ワーグナー アレックス, Alex Y. Wagner
フィン ラセル, Finn Edgar Russel
```

Then run the script and it will output a file called `member_names.csv` which can be imported into an Excel sheet and used for the VB script described above.


