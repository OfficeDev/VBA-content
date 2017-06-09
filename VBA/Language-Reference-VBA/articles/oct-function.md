---
title: Oct Function
keywords: vblr6.chm1008983
f1_keywords:
- vblr6.chm1008983
ms.prod: office
ms.assetid: 178a6099-9181-2160-2b97-e08c97f8b2bb
ms.date: 06/08/2017
---


# Oct Function



Returns a  **Variant** ( **String** ) representing the octal value of a number.
 **Syntax**
 **Oct** ( _number_ )
The required  _number_[argument](vbe-glossary.md) is any valid[numeric expression](vbe-glossary.md) or[string expression](vbe-glossary.md).
 **Remarks**
If  _number_ is not already a whole number, it is rounded to the nearest whole number before being evaluated.


|**If  _number_ is**|**Oct returns**|
|:-----|:-----|
|[Null](vbe-glossary.md)|**Null**|
|[Empty](vbe-glossary.md)|Zero (0)|
|Any other number|Up to 11 octal characters|
You can represent octal numbers directly by preceding numbers in the proper range with  `&;O`. For example, . For example,  `&;O10` is the octal notation for decimal 8.

## Example

This example uses the  **Oct** function to return the octal value of a number.


```vb
Dim MyOct
MyOct = Oct(4)     ' Returns 4.
MyOct = Oct(8)    ' Returns 10.
MyOct = Oct(459)    ' Returns 713.


```


