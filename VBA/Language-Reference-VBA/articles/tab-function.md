---
title: Tab Function
keywords: vblr6.chm1009039
f1_keywords:
- vblr6.chm1009039
ms.prod: office
ms.assetid: 609036b5-08c8-fb5c-4959-3e1a4e108f8d
ms.date: 06/08/2017
---


# Tab Function



Used with the  **Print #** statement or the **Print** method to position output.
 **Syntax**
 **Tab** [ **(**_n_**)** ]
The optional  _n_[argument](vbe-glossary.md) is the column number moved to before displaying or printing the next[expression](vbe-glossary.md) in a list. If omitted, **Tab** moves the insertion point to the beginning of the next[print zone](vbe-glossary.md). This allows  **Tab** to be used instead of a comma in[locales](vbe-glossary.md) where the comma is used as a decimal separator.
 **Remarks**
If the current print position on the current line is greater than  _n_, **Tab** skips to the _n_ th column on the next output line. If _n_ is less than 1, **Tab** moves the print position to column 1. If _n_ is greater than the output line width, **Tab** calculates the next print position using the formula:
 _n_**Mod**_width_
For example, if  _width_ is 80 and you specify **Tab(** 90 **)**, the next print will start at column 10 (the remainder of 90/80). If _n_ is less than the current print position, printing begins on the next line at the calculated print position. If the calculated print position is greater than the current print position, printing begins at the calculated print position on the same line.
The leftmost print position on an output line is always 1. When you use the  **Print #** statement to print to files, the rightmost print position is the current width of the output file, which you can set using the **Width #** statement.

 **Note**  Make sure your tabular columns are wide enough to accommodate wide letters.

When you use the  **Tab** function with the **Print** method, the print surface is divided into uniform, fixed-width columns. The width of each column is an average of the width of all characters in the point size for the chosen font. However, there is no correlation between the number of characters printed and the number of fixed-width columns those characters occupy. For example, the uppercase letter W occupies more than one fixed-width column and the lowercase letter i occupies less than one fixed-width column.

## Example

This example uses the  **Tab** function to position output in a file and in the **Immediate** window.


```vb
' The Tab function can be used with the Print # statement.
Open "TESTFILE" For Output As #1    ' Open file for output.
' The second word prints at column 20.
Print #1, "Hello"; Tab(20); "World."
' If the argument is omitted, cursor is moved to the next print zone.
Print #1, "Hello"; Tab; "World"
Close #1    ' Close file.

```

The  **Tab** function can also be used with the **Print** method. The following statement prints text starting at column 10.




```vb
Debug.Print Tab(10); "10 columns from start."


```


