---
title: Or Operator
keywords: vblr6.chm1008993
f1_keywords:
- vblr6.chm1008993
ms.prod: office
ms.assetid: 3b0e4886-2f84-1296-9428-69338d033c6c
ms.date: 06/08/2017
---


# Or Operator



Used to perform a logical disjunction on two [expressions](vbe-glossary.md).
 **Syntax**
 _result_**=**_expression1_ **Or** _expression2_
The  **Or** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](vbe-glossary.md).|
| _expression1_|Required; any expression.|
| _expression2_|Required; any expression.|
 **Remarks**
If either or both expressions evaluate to  **True**, _result_ is **True**. The following table illustrates how _result_ is determined:


|**If  _expression1_ is**|**And  _expression2_ is**|**Then  _result_ is**|
|:-----|:-----|:-----|
|**True**|**True**|**True**|
|**True**|**False**|**True**|
|**True**|[Null](vbe-glossary.md)|**True**|
|**False**|**True**|**True**|
|**False**|**False**|**False**|
|**False**|**Null**|**Null**|
|**Null**|**True**|**True**|
|**Null**|**False**|**Null**|
|**Null**|**Null**|**Null**|
The  **Or** operator also performs a [bitwise comparison](vbe-glossary.md) of identically positioned bits in two [numeric expressions](vbe-glossary.md) and sets the corresponding bit in _result_ according to the following table:


|**If bit in  _expression1_ is**|**And bit in  _expression2_ is**|**Then  _result_ is**|
|:-----|:-----|:-----|
|0|0|0|
|0|1|1|
|1|0|1|
|1|1|1|

## Example

This example uses the  **Or** operator to perform logical disjunction on two expressions.


```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = A > B Or B > C    ' Returns True.
MyCheck = B > A Or B > C    ' Returns True.
MyCheck = A > B Or B > D    ' Returns True.
MyCheck = B > D Or B > A    ' Returns Null.
MyCheck = A Or B    ' Returns 10 (bitwise comparison).


```


