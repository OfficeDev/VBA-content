---
title: LTrim, RTrim, and Trim Functions
keywords: vblr6.chm1008797
f1_keywords:
- vblr6.chm1008797
ms.prod: office
ms.assetid: ffe13d6f-8e7a-3413-98a1-3263c771178b
ms.date: 06/08/2017
---


# LTrim, RTrim, and Trim Functions



Returns a  **Variant** ( **String** ) containing a copy of a specified string without leading spaces ( **LTrim** ), trailing spaces ( **RTrim** ), or both leading and trailing spaces ( **Trim** ).
 **Syntax**
 **LTrim** ( _string_ )
 **RTrim** ( _string_ )
 **Trim** ( _string_ )
The required  _string_ [argument](vbe-glossary.md) is any valid [string expression](vbe-glossary.md). If  _string_ contains [Null](vbe-glossary.md),  **Null** is returned.

## Example

This example uses the  **LTrim** function to strip leading spaces and the **RTrim** function to strip trailing spaces from a string variable. It uses the **Trim** function to strip both types of spaces.


```vb
Dim MyString, TrimString
MyString = "  <-Trim->  "    ' Initialize string.
TrimString = LTrim(MyString)    ' TrimString = "<-Trim->  ".
TrimString = RTrim(MyString)    ' TrimString = "  <-Trim->".
TrimString = LTrim(RTrim(MyString))    ' TrimString = "<-Trim->".
' Using the Trim function alone achieves the same result.
TrimString = Trim(MyString)    ' TrimString = "<-Trim->".


```


