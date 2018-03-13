---
title: Imp Operator
keywords: vblr6.chm1008941
f1_keywords:
- vblr6.chm1008941
ms.prod: office
ms.assetid: 7f1d82c0-de89-40ae-a504-804d7cf04e28
ms.date: 06/08/2017
---


# Imp Operator



Used to perform a logical implication on two [expressions](vbe-glossary.md).
 **Syntax**
 _result_**=**_expression1_ **Imp** _expression2_.
The  **Imp** operator syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                       |
|:----------------------|:---------------------------------------------------|
| <em>result</em>       | Required; any numeric [variable](vbe-glossary.md). |
| <em>expression1</em>  | Required; any expression.                          |
| <em>expression2</em>  | Required; any expression.                          |

 **Remarks**
The following table illustrates how  _result_ is determined:


| <strong>If  <em>expression1</em> is</strong> | <strong>And  <em>expression2</em> is</strong> | <strong>The  <em>result</em> is</strong> |
|:---------------------------------------------|:----------------------------------------------|:-----------------------------------------|
| <strong>True</strong>                        | <strong>True</strong>                         | <strong>True</strong>                    |
| <strong>True</strong>                        | <strong>False</strong>                        | <strong>False</strong>                   |
| <strong>True</strong>                        | [Null](vbe-glossary.md)                       | <strong>Null</strong>                    |
| <strong>False</strong>                       | <strong>True</strong>                         | <strong>True</strong>                    |
| <strong>False</strong>                       | <strong>False</strong>                        | <strong>True</strong>                    |
| <strong>False</strong>                       | <strong>Null</strong>                         | <strong>True</strong>                    |
| <strong>Null</strong>                        | <strong>True</strong>                         | <strong>True</strong>                    |
| <strong>Null</strong>                        | <strong>False</strong>                        | <strong>Null</strong>                    |
| <strong>Null</strong>                        | <strong>Null</strong>                         | <strong>Null</strong>                    |

The  **Imp** operator performs a [bitwise comparison](vbe-glossary.md) of identically positioned bits in two [numeric expressions](vbe-glossary.md) and sets the corresponding bit in _result_ according to the following table:


|**If bit in  _expression1_ is**|**And bit in  _expression2_ is**|**The  _result_ is**|
|:-----|:-----|:-----|
|0|0|1|
|0|1|1|
|1|0|0|
|1|1|1|

## Example

This example uses the  **Imp** operator to perform logical implication on two expressions.


```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = A > B Imp B > C    ' Returns True.
MyCheck = A > B Imp C > B    ' Returns False.
MyCheck = B > A Imp C > B    ' Returns True.
MyCheck = B > A Imp C > D    ' Returns True.
MyCheck = C > D Imp B > A    ' Returns Null.
MyCheck = B Imp A    ' Returns -1 (bitwise comparison).
```


