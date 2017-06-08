---
title: LSet Statement
keywords: vblr6.chm1008969
f1_keywords:
- vblr6.chm1008969
ms.prod: office
ms.assetid: ecf1dbcb-7f8d-9f98-2d49-f7ceb790415d
ms.date: 06/08/2017
---


# LSet Statement

Left aligns a string within a string [variable](vbe-glossary.md), or copies a variable of one [user-defined type](vbe-glossary.md) to another variable of a different user-defined type.

 **Syntax**

 **LSet**_stringvar_**=**_string_

 **LSet**_varname1_**=**_varname2_
The  **LSet** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _stringvar_|Required. Name of string [variable](vbe-glossary.md).|
| _string_|Required. [String expression](vbe-glossary.md) to be left-aligned within _stringvar._|
| _varname1_|Required. Variable name of the user-defined type being copied to.|
| _varname2_|Required. Variable name of the user-defined type being copied from.|
 **Remarks**
 **LSet** replaces any leftover characters in _stringvar_ with spaces.
If  _string_ is longer than _stringvar_, **LSet** places only the leftmost characters, up to the length of the _stringvar,_ in _stringvar_.
Using  **LSet** to copy a variable of one user-defined type into a variable of a different user-defined type is not recommended. Copying data of one[data type](vbe-glossary.md) into space reserved for a different data type can cause unpredictable results.
When you copy a variable from one user-defined type to another, the binary data from one variable is copied into the memory space of the other, without regard for the data types specified for the elements.

## Example

This example uses the  **LSet** statement to left align a string within a string variable. Although **LSet** can also be used to copy a variable of one user-defined type to another variable of a different, but compatible, user-defined type, this practice is not recommended. Due to the varying implementations of data structures among platforms, such a use of **LSet** can't be guaranteed to be portable.


```vb
Dim MyString 
MyString = "0123456789" ' Initialize string. 
Lset MyString = "<-Left" ' MyString contains "<-Left ". 

```


