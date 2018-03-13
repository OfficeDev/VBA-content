---
title: Mid Statement
keywords: vblr6.chm1011353
f1_keywords:
- vblr6.chm1011353
ms.prod: office
ms.assetid: a9923853-55d5-5b50-d422-57cba84d9f47
ms.date: 06/08/2017
---


# Mid Statement

Replaces a specified number of characters in a  **Variant** ( **String** )[variable](vbe-glossary.md) with characters from another string.

 **Syntax**

 **Mid** ( _stringvar_, _start_ [, _length_ ]) **=**_string_

The  **Mid** statement syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                  |
|:----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------|
| <em>stringvar</em>    | Required. Name of string variable to modify.                                                                                                  |
| <em>start</em>        | Required;  <strong>Variant</strong> ( <strong>Long</strong> ). Character position in <em>stringvar</em> where the replacement of text begins. |
| <em>length</em>       | Optional;  <strong>Variant</strong> ( <strong>Long</strong> ). Number of characters to replace. If omitted, all of <em>string</em> is used.   |
| <em>string</em>       | Required. [String expression](vbe-glossary.md) that replaces part of <em>stringvar</em>.                                                      |

 **Remarks**
The number of characters replaced is always less than or equal to the number of characters in  _stringvar_.

 **Note**  Use the  **MidB** statement with byte data contained in a string. In the **MidB** statement, _start_ specifies the byte position within _stringvar_ where replacement begins and _length_ specifies the numbers of bytes to replace.


## Example

This example uses the  **Mid** statement to replace a specified number of characters in a string variable with characters from another string.


```vb
Dim MyString 
MyString = "The dog jumps" ' Initialize string. 
Mid(MyString, 5, 3) = "fox" ' MyString = "The fox jumps". 
Mid(MyString, 5) = "cow" ' MyString = "The cow jumps". 
Mid(MyString, 5) = "cow jumped over" ' MyString = "The cow jumpe". 
Mid(MyString, 5, 3) = "duck" ' MyString = "The duc jumpe". 
```


