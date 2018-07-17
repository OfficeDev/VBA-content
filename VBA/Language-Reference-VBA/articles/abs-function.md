---
title: Abs Function
keywords: vblr6.chm1008850
f1_keywords:
- vblr6.chm1008850
ms.prod: office
ms.assetid: b2184f54-bf2b-a3da-f1c8-b38575a213eb
ms.date: 06/08/2017
---


# Abs Function


Returns a value of the same type that is passed to it specifying the absolute value of a number.

## Syntax
**Abs(**_number_**)**
 
The required _number_ [argument](vbe-glossary.md) can be any valid[numeric expression](vbe-glossary.md). If _number_ contains [Null](vbe-glossary.md), **Null** is returned; if it is an uninitialized [variable](vbe-glossary.md), zero is returned.

### Remarks
The absolute value of a number is its unsigned magnitude. For example,  `ABS(-1)` and `ABS(1)` both return `1`.

## Example

This example uses the  **Abs** function to compute the absolute value of a number.


```vb
Dim MyNumber
MyNumber = Abs(50.3)    ' Returns 50.3.
MyNumber = Abs(-50.3)    ' Returns 50.3.
```


