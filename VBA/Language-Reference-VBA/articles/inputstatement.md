---
title: Input  Statement
keywords: vblr6.chm1008943
f1_keywords:
- vblr6.chm1008943
ms.prod: office
ms.assetid: b248ddce-f733-8bb2-2bea-349f5d2c6552
ms.date: 06/08/2017
---


# Input # Statement

Reads data from an open sequential file and assigns the data to [variables](vbe-glossary.md).

 **Syntax**

 **Input** **#**_filenumber, varlist_

The  **Input #** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](vbe-glossary.md).|
| _varlist_|Required. Comma-delimited list of variables that are assigned values read from the file — can't be an [array](vbe-glossary.md) or[object variable](vbe-glossary.md). However, variables that describe an element of an array or [user-defined type](vbe-glossary.md) may be used.|
 **Remarks**
Data read with  **Input #** is usually written to a file with **Write #**. Use this[statement](vbe-glossary.md) only with files opened in **Input** or **Binary** mode.
When read, standard string or numeric data is assigned to variables without modification. The following table illustrates how other input data is treated:


|**Data**|**Value assigned to variable**|
|:-----|:-----|
|Delimiting comma or blank line|[Empty](vbe-glossary.md)|
|#NULL#|[Null](vbe-glossary.md)|
|#TRUE# or #FALSE#|**True** or **False**|
|# _yyyy-mm-dd hh:mm:ss_ #|The date and/or time represented by the [expression](vbe-glossary.md)|
|#ERROR  _errornumber_ #| _errornumber_ (variable is a[Variant](vbe-glossary.md) tagged as an error)|
Double quotation marks () within input data are ignored.

 **Note**  You should not write strings that contain embedded quotation marks, for example,  `"1,2""X"` for use with the **Input #** statement: **Input #** parses this string as two complete and separate strings.

Data items in a file must appear in the same order as the variables in  _varlist_ and match variables of the same[data type](vbe-glossary.md). If a variable is numeric and the data is not numeric, a value of zero is assigned to the variable.
If you reach the end of the file while you are inputting a data item, the input is terminated and an error occurs.

 **Note**  To be able to correctly read data from a file into variables using  **Input #**, use the **Write #** statement instead of the **Print #** statement to write the data to the files. Using **Write #** ensures each separate data field is properly delimited.


## Example

This example uses the  **Input #** statement to read data from a file into two variables. This example assumes that is a file with a few lines of data written to it using the **Write #** statement; that is, each line contains a string in quotations and a number separated by a comma, for example, ("Hello", 234).


```vb
Dim MyString, MyNumber 
Open "TESTFILE" For Input As #1    ' Open file for input. 
Do While Not EOF(1)    ' Loop until end of file. 
    Input #1, MyString, MyNumber    ' Read data into two variables. 
    Debug.Print MyString, MyNumber    ' Print data to the Immediate window. 
Loop 
Close #1    ' Close file. 

```


