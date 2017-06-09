---
title: And Operator
keywords: vblr6.chm1008852
f1_keywords:
- vblr6.chm1008852
ms.prod: office
ms.assetid: 523e8cd3-f27c-2ec5-62e8-e95686a9f9ac
ms.date: 06/08/2017
---


# And Operator



Used to perform a logical conjunction on two [expressions](vbe-glossary.md).
 **Syntax**
 _result_**=**_expression1_ **And** _expression2_
The  **And** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](vbe-glossary.md).|
| _expression1_|Required; any expression.|
| _expression2_|Required; any expression.|
 **Remarks**
If both expressions evaluate to  **True**, _result_ is **True**. If either expression evaluates to **False**, _result_ is **False.** The following table illustrates how _result_ is determined:


|**If  _expression1_ is**|**And  _expression2_ is**|**The  _result_ is**|
|:-----|:-----|:-----|
|**True**|**True**|**True**|
|**True**|**False**|**False**|
|**True**|[Null](vbe-glossary.md)|**Null**|
|**False**|**True**|**False**|
|**False**|**False**|**False**|
|**False**|**Null**|**False**|
|**Null**|**True**|**Null**|
|**Null**|**False**|**False**|
|**Null**|**Null**|**Null**|
The  **And** operator also performs a [bitwise comparison](vbe-glossary.md) of identically positioned bits in two [numeric expressions](vbe-glossary.md) and sets the corresponding bit in _result_ according to the following table:


|**If bit in  _expression1_ is**|**And bit in  _expression2_ is**|**The  _result_ is**|
|:-----|:-----|:-----|
|0|0|0|
|0|1|0|
|1|0|0|
|1|1|1|

## Example

This example uses the  **And** operator to perform a logical conjunction on two expressions.


```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = A > B And B > C    ' Returns True.
MyCheck = B > A And B > C    ' Returns False.
MyCheck = A > B And B > D    ' Returns Null.
MyCheck = A And B    ' Returns 8 (bitwise comparison).

```


