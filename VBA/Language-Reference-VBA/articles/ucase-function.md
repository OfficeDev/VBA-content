---
title: UCase Function
keywords: vblr6.chm1009051
f1_keywords:
- vblr6.chm1009051
ms.prod: office
ms.assetid: 444bd68b-a2bf-11b2-e6b7-76edf9b03ecd
ms.date: 06/08/2017
---


# UCase Function



Returns a  **Variant** ( **String** ) containing the specified string, converted to uppercase.
 **Syntax**
 **UCase** ( _string_ )
The required  _string_[argument](vbe-glossary.md) is any valid[string expression](vbe-glossary.md). If  _string_ contains[Null](vbe-glossary.md),  **Null** is returned.
 **Remarks**
Only lowercase letters are converted to uppercase; all uppercase letters and nonletter characters remain unchanged.

## Example

This example uses the  **UCase** function to return an uppercase version of a string.


```vb
Dim LowerCase, UpperCase
LowerCase = "Hello World 1234"    ' String to convert.
UpperCase = UCase(LowerCase)    ' Returns "HELLO WORLD 1234".


```


