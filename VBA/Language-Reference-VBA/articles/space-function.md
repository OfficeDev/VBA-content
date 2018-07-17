---
title: Space Function
keywords: vblr6.chm1009026
f1_keywords:
- vblr6.chm1009026
ms.prod: office
ms.assetid: fa531cfb-863f-ede9-34b8-6000711d71ed
ms.date: 06/08/2017
---


# Space Function



Returns a  **Variant** ( **String** ) consisting of the specified number of spaces.
 **Syntax**
 **Space** ( _number_ )
The required  _number_[argument](vbe-glossary.md) is the number of spaces you want in the string.
 **Remarks**
The  **Space** function is useful for formatting output and clearing data in fixed-length strings.

## Example

This example uses the  **Space** function to return a string consisting of a specified number of spaces.


```vb
Dim MyString
' Returns a string with 10 spaces.
MyString = Space(10)

' Insert 10 spaces between two strings.
MyString = "Hello" &; Space(10) &; "World"


```


