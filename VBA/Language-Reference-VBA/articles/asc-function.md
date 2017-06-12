---
title: Asc Function
keywords: vblr6.chm1009247
f1_keywords:
- vblr6.chm1009247
ms.prod: office
ms.assetid: 4c5775f4-792f-f9d0-6eff-41d6fff9048c
ms.date: 06/08/2017
---


# Asc Function



Returns an [Integer](vbe-glossary.md) representing the[character code](vbe-glossary.md) corresponding to the first letter in a string.
 **Syntax**
 **Asc(**_string_**)**
The required  _string_[argument](vbe-glossary.md) is any valid[string expression](vbe-glossary.md). If the  _string_ contains no characters, a[run-time error](vbe-glossary.md) occurs.
 **Remarks**
The range for returns is 0 - 255 on non-DBCS systems, but -32768 - 32767 on [DBCS](vbe-glossary.md) systems.

 **Note**  The  **AscB** function is used with byte data contained in a string. Instead of returning the character code for the first character, **AscB** returns the first byte. The **AscW** function returns the[Unicode](vbe-glossary.md) character code except on platforms where Unicode is not supported, in which case, the behavior is identical to the **Asc** function.


 **Note**  Visual Basic for the Macintosh does not support Unicode strings. Therefore,  **AscW** ( _n_ ) cannot return all Unicode characters for n values in the range of 128 - 65,535, as it does in the Windows environment. Instead, **AscW** ( _n_ ) attempts a "best guess" for Unicode values n greater than 127. Therefore, you should not use **AscW** in the Macintosh environment.


## Example

This example uses the  **Asc** function to return a character code corresponding to the first letter in the string.


```vb
Dim MyNumber
MyNumber = Asc("A")    ' Returns 65.
MyNumber = Asc("a")    ' Returns 97.
MyNumber = Asc("Apple")    ' Returns 65.


```


