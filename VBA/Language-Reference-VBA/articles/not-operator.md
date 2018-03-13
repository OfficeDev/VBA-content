---
title: Not Operator
keywords: vblr6.chm1008981
f1_keywords:
- vblr6.chm1008981
ms.prod: office
ms.assetid: e5ae5a73-4f34-0071-ee67-98e4ca519748
ms.date: 06/08/2017
---


# Not Operator



Used to perform logical negation on an [expression](vbe-glossary.md).
 **Syntax**
 _result_**=** **Not**_expression_
The  **Not** operator syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                       |
|:----------------------|:---------------------------------------------------|
| <em>result</em>       | Required; any numeric [variable](vbe-glossary.md). |
| <em>expression</em>   | Required; any expression.                          |

 **Remarks**
The following table illustrates how  _result_ is determined:


| <strong>If  <em>expression</em> is</strong> | <strong>Then  <em>result</em> is</strong> |
|:--------------------------------------------|:------------------------------------------|
| <strong>True</strong>                       | <strong>False</strong>                    |
| <strong>False</strong>                      | <strong>True</strong>                     |
| [Null](vbe-glossary.md)                     | <strong>Null</strong>                     |

In addition, the  **Not** operator inverts the bit values of any variable and sets the corresponding bit in _result_ according to the following table:


|**If bit in  _expression_ is**|**Then bit in  _result_ is**|
|:-----|:-----|
|o|1|
|1|0|

## Example

This example uses the  **Not** operator to perform logical negation on an expression.


```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = Not(A > B)    ' Returns False.
MyCheck = Not(B > A)    ' Returns True.
MyCheck = Not(C > D)    ' Returns Null.
MyCheck = Not A    ' Returns -11 (bitwise comparison).
```


