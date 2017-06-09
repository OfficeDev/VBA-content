---
title: Str Function
keywords: vblr6.chm1011369
f1_keywords:
- vblr6.chm1011369
ms.prod: office
ms.assetid: bb9c4e8c-c3ea-5021-aa4c-473e30b64902
ms.date: 06/08/2017
---


# Str Function



Returns a  **Variant** ( **String** ) representation of a number.
 **Syntax**
 **Str** ( _number_ )
The required  _number_[argument](vbe-glossary.md) is a[Long](vbe-glossary.md) containing any valid[numeric expression](vbe-glossary.md).
 **Remarks**
When numbers are converted to strings, a leading space is always reserved for the sign of  _number_. If _number_ is positive, the returned string contains a leading space and the plus sign is implied.
Use the  **Format** function to convert numeric values you want formatted as dates, times, or currency or in other user-defined formats. Unlike **Str**, the **Format** function doesn't include a leading space for the sign of _number_.

 **Note**  The  **Str** function recognizes only the period ( **.** ) as a valid decimal separator. When different decimal separators may be used (for example, in international applications), use **CStr** to convert a number to a string.


## Example

This example uses the  **Str** function to return a string representation of a number. When a number is converted to a string, a leading space is always reserved for its sign.


```vb
Dim MyString
MyString = Str(459)    ' Returns " 459".
MyString = Str(-459.65)    ' Returns "-459.65".
MyString = Str(459.001)    ' Returns " 459.001".


```


