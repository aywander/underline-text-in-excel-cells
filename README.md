# underline-text-in-excel-cells

A VB script (`underline_substrings.vb`) to automatically underline sub-strings in a range of Excel cells and a python script (`create_names_csv.py`) to create variations of name strings.

## VB script to underline sub-strings in Excel cells

In Excel, it is not possible to underline (or change formatting) of a sub-string in a cell with search-and-replace, conditioanl formatting or formulae.
A workaround is to copy text to a word document do search-and-replace with formatting options, but the VB script in this repository provides a fully programatic one-click solution.

Sometimes, we need to underline specific sub-strings, like names, in an excel data sheet for administrative purposes. The VB script `underline_substrings.vb` was conceived to automate this arduous task.

The VB script can be used not just for names, but for any other strings, and other formatting changes within cells can of course also be performed.

### Preparation

Adjust the sheet name and range in the VB script to match the cells that contain all the text to search for underlining.

```vb
Dim NamesVector
NamesVector = Create_Vector(Sheets("Sheet2").Range("a1:b4"))
```

In the example above, cells `a1:b4` in sheet `Sheet2` each contain text to search for and underline in the current selection.

### Execution

Select a range of cells in the data where underlining should occur and run the script in the VBA editor of the Excel file.

### Modifications

To change or add formatting that should be applied, change the following line or add appropriate lines below this line:

```
cl.Characters(StartPos, TotalLen).Font.Underline = xlUnderlineStyleSingle
```

See [the following page](https://docs.microsoft.com/en-us/office/vba/api/excel.font(object)) for properties of the `Characters.Font` object.

## Python helper script to generate variations on name strings

The python script `create_names_csv.py` is somewhat tailored for administrative tasks at the CCS, and maybe also other institutes, but could find broader use.

This script reads a simple two-column text file of member names and outputs a larger variety of name arrangements in csv format, which can be used with the VB script above to underline these names in an Excel sheet.

The two-column text file should be named `member_names.txt` and be located in the same directory as the python script. The two-column data is separated by a comma. Left column: kanji/kana name `<Surname> <First name>` (space between surname and first name). Right column: romaji name `<First names> <Surname>` (space between first name and surname). Middle names possible, but only for romaji name.

E.g.
```txt
佐藤 健二, Kenji Sato
ワーグナー アレックス, Alex Y. Wagner
フィン ラセル, Finn Edgar Russel
```

Then run the script and it will output a file called `member_names.csv` which can be imported into an Excel sheet and used for the VB script described above.
When importing the file into Excel, be sure to set the options to delimited, delimted by comma, and the encoding to UTF-8.

