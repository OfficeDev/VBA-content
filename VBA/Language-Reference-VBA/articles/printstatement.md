---
title: Print  Statement
keywords: vblr6.chm1008995
f1_keywords:
- vblr6.chm1008995
ms.prod: office
ms.assetid: 47c69cf9-2476-b9c2-782c-1c0fc2747936
ms.date: 06/08/2017
---


# Print # Statement

Writes display-formatted data to a sequential file.

 **Syntax**

 **Print** **#**_filenumber_, [ _outputlist_ ]

The  **Print #** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](vbe-glossary.md).|
| _outputlist_|Optional. [Expression](vbe-glossary.md) or list of expressions to print.|
 **Settings**
The  _outputlist_[argument](vbe-glossary.md) settings are:
[{ **Spc(**_n_**)** |**Tab** [ **(**_n_**)** ]}] [ _expression_ ] [ _charpos_ ]


|**Setting**|**Description**|
|:-----|:-----|
|**Spc(**_n_**)**|Used to insert space characters in the output, where  _n_ is the number of space characters to insert.|
|**Tab(**_n_**)**|Used to position the insertion point to an absolute column number, where  _n_ is the column number. Use **Tab** with no argument to position the insertion point at the beginning of the next[print zone](vbe-glossary.md).|
| _expression_|[Numeric expressions](vbe-glossary.md) or[string expressions](vbe-glossary.md) to print.|
| _charpos_|Specifies the insertion point for the next character. Use a semicolon to position the insertion point immediately after the last character displayed. Use  **Tab(**_n_**)** to position the insertion point to an absolute column number. Use **Tab** with no argument to position the insertion point at the beginning of the next print zone. If _charpos_ is omitted, the next character is printed on the next line.|
 **Remarks**
Data written with  **Print #** is usually read from a file with **Line Input #** or **Input**.
If you omit  _outputlist_ and include only a list separator after _filenumber_, a blank line is printed to the file. Multiple expressions can be separated with either a space or a semicolon. A space has the same effect as a semicolon.
For [Boolean](vbe-glossary.md) data, either `True` or or `False` is printed. The **True** and **False** keywords are not translated, regardless of the[locale](vbe-glossary.md).
[Date](vbe-glossary.md) data is written to the file using the standard short date format recognized by your system. When either the date or the time component is missing or zero, only the part provided gets written to the file.
Nothing is written to the file if  _outputlist_ data is[Empty](vbe-glossary.md). However, if  _outputlist_ data is[Null](vbe-glossary.md),  **Null** is written to the file.
For  **Error** data, the output appears as `Error` _errorcode_. The **Error** keyword is not translated regardless of the locale.
All data written to the file using  **Print #** is internationally aware; that is, the data is properly formatted using the appropriate decimal separator.
Because  **Print #** writes an image of the data to the file, you must delimit the data so it prints correctly. If you use **Tab** with no arguments to move the print position to the next print zone, **Print #** also writes the spaces between print fields to the file.

 **Note**  If, at some future time, you want to read the data from a file using the  **Input #** statement, use the **Write #** statement instead of the **Print #** statement to write the data to the file. Using **Write #** ensures the integrity of each separate data field by properly delimiting it, so it can be read back in using **Input #**. Using **Write #** also ensures it can be correctly read in any locale.


## Example

This example uses the  **Print #** statement to write data to a file.


```vb
Open "TESTFILE" For Output As #1 ' Open file for output. 
Print #1, "This is a test" ' Print text to file. 
Print #1, ' Print blank line to file. 
Print #1, "Zone 1"; Tab ; "Zone 2" ' Print in two print zones. 
Print #1, "Hello" ; " " ; "World" ' Separate strings with space. 
Print #1, Spc(5) ; "5 leading spaces " ' Print five leading spaces. 
Print #1, Tab(10) ; "Hello" ' Print word at column 10. 
 
' Assign Boolean, Date, Null and Error values. 
Dim MyBool, MyDate, MyNull, MyError 
MyBool = False : MyDate = #February 12, 1969# : MyNull = Null 
MyError = CVErr(32767) 
' True, False, Null, and Error are translated using locale settings of 
' your system. Date literals are written using standard short date 
' format. 
Print #1, MyBool ; " is a Boolean value" 
Print #1, MyDate ; " is a date" 
Print #1, MyNull ; " is a null value" 
Print #1, MyError ; " is an error value" 
Close #1 ' Close file. 

```


