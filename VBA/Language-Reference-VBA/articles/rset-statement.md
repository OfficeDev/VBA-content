---
title: RSet Statement
keywords: vblr6.chm1009009
f1_keywords:
- vblr6.chm1009009
ms.prod: office
ms.assetid: 07a4f730-ef85-cbeb-30ac-ea51d161f27f
ms.date: 06/08/2017
---


# RSet Statement

Right aligns a string within a string [variable](vbe-glossary.md).

 **Syntax**

 **RSet**_stringvar_**=**_string_

The  **RSet** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _stringvar_|Required. Name of string variable.|
| _string_|Required. [String expression](vbe-glossary.md) to be right-aligned within _stringvar_.|
 **Remarks**
If  _stringvar_ is longer than _string_, **RSet** replaces any leftover characters in _stringvar_ with spaces, back to its beginning.

 **Note**   **RSet** can't be used with[user-defined types](vbe-glossary.md).


## Example

This example uses the  **RSet** statement to right align a string within a string variable.


```vb
Dim MyString 
MyString = "0123456789" ' Initialize string. 
Rset MyString = "Right->" ' MyString contains " Right->". 

```


