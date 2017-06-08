---
title: String Function
keywords: vblr6.chm1011358
f1_keywords:
- vblr6.chm1011358
ms.prod: office
ms.assetid: d6c5c054-21b9-f777-acae-ac31710ba5c5
ms.date: 06/08/2017
---


# String Function



Returns a  **Variant** ( **String** ) containing a repeating character string of the length specified.
 **Syntax**
 **String** ( **_number_**, **_character_** )
The  **String** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_number_**|Required; [Long](vbe-glossary.md). Length of the returned string. If  **_number_** contains[Null](vbe-glossary.md),  **Null** is returned.|
|**_character_**|Required; [Variant](vbe-glossary.md). [Character code](vbe-glossary.md) specifying the character or[string expression](vbe-glossary.md) whose first character is used to build the return string. If **_character_** contains **Null**, **Null** is returned.|
 **Remarks**
If you specify a number for  **_character_** greater than 255, **String** converts the number to a valid character code using the formula:
 **_character_** **Mod** 256

## Example

This example uses the  **String** function to return repeating character strings of the length specified.


```vb
Dim MyString
MyString = String(5, "*")    ' Returns "*****".
MyString = String(5, 42)    ' Returns "*****".
MyString = String(10, "ABC")    ' Returns "AAAAAAAAAA".


```


