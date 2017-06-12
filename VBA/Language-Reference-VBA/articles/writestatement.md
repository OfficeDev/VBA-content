---
title: Write  Statement
keywords: vblr6.chm1009061
f1_keywords:
- vblr6.chm1009061
ms.prod: office
ms.assetid: b39df18a-4cdc-2aca-d941-35cffe8d0005
ms.date: 06/08/2017
---


# Write # Statement

Writes data to a sequential file.

 **Syntax**

 **Write #**_filenumber_, [ _outputlist_ ]

The  **Write #** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](vbe-glossary.md).|
| _outputlist_|Optional. One or more comma-delimited [numeric expressions](vbe-glossary.md) or[string expressions](vbe-glossary.md) to write to a file.|
 **Remarks**
Data written with  **Write #** is usually read from a file with **Input #**.
If you omit  _outputlist_ and include a comma after _filenumber_, a blank line is printed to the file. Multiple expressions can be separated with a space, a semicolon, or a comma. A space has the same effect as a semicolon.
When  **Write #** is used to write data to a file, several universal assumptions are followed so the data can always be read and correctly interpreted using **Input #**, regardless of[locale](vbe-glossary.md):


- Numeric data is always written using the period as the decimal separator.
    
- For [Boolean](vbe-glossary.md) data, either `#TRUE#` or `#FALSE#` is printed. The **True** and **False**[keywords](vbe-glossary.md) are not translated, regardless of locale.
    
- [Date](vbe-glossary.md) data is written to the file using the[universal date format](vbe-glossary.md). When either the date or the time component is missing or zero, only the part provided gets written to the file.
    
- Nothing is written to the file if  _outputlist_ data is[Empty](vbe-glossary.md). However, for [Null](vbe-glossary.md) data, `#NULL#` is written.
    
- If  _outputlist_ data is **Null** data, `#NULL#` is written to the file.
    
- For  **Error** data, the output appears as `#ERROR errorcode#`. The  **Error** keyword is not translated, regardless of locale.
    

Unlike the  **Print #** statement, the **Write #** statement inserts commas between items and quotation marks around strings as they are written to the file. You don't have to put explicit delimiters in the list. **Write #** inserts a newline character, that is, a carriage return-linefeed ( **Chr(** 13 **)** + **Chr(** 10 **)** ), after it has written the final character in _outputlist_ to the file.

 **Note**  You should not write strings that contain embedded quotation marks, for example, `"1,2""X"` for use with the **Input #** statement: **Input #** parses this string as two complete and separate strings.


## Example

This example uses the  **Write #** statement to write raw data to a sequential file.


```
Open "TESTFILE" For Output As #1    ' Open file for output. 
Write #1, "Hello World", 234    ' Write comma-delimited data. 
Write #1,    ' Write blank line. 
 
Dim MyBool, MyDate, MyNull, MyError 
' Assign Boolean, Date, Null, and Error values. 
MyBool = False : MyDate = #February 12, 1969# : MyNull = Null 
MyError = CVErr(32767) 
' Boolean data is written as #TRUE# or #FALSE#. Date literals are  
' written in universal date format, for example, #1994-07-13#  
 'represents July 13, 1994. Null data is written as #NULL#.  
' Error data is written as #ERROR errorcode#. 
Write #1, MyBool ; " is a Boolean value" 
Write #1, MyDate ; " is a date" 
Write #1, MyNull ; " is a null value" 
Write #1, MyError ; " is an error value" 
Close #1    ' Close file. 

```


