---
title: Xor Operator
keywords: vblr6.chm1009062
f1_keywords:
- vblr6.chm1009062
ms.prod: office
ms.assetid: 30f2f390-e777-8793-a287-038fb9a18ce6
ms.date: 06/08/2017
---


# Xor Operator



Used to perform a logical exclusion on two [expressions](vbe-glossary.md).
 **Syntax**
[ _result_**=** ] _expression1_ **Xor** _expression2_
The  **Xor** operator syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                       |
|:----------------------|:---------------------------------------------------|
| <em>result</em>       | Optional; any numeric [variable](vbe-glossary.md). |
| <em>expression1</em>  | Required; any expression.                          |
| <em>expression2</em>  | Required; any expression.                          |

 **Remarks**
If one, and only one, of the expressions evaluates to  **True**, _result_ is **True**. However, if either expression is [Null](vbe-glossary.md),  _result_ is also **Null**. When neither expression is **Null**, _result_ is determined according to the following table:


| <strong>If  <em>expression1</em> is</strong> | <strong>And  <em>expression2</em> is</strong> | <strong>Then  <em>result</em> is</strong> |
|:---------------------------------------------|:----------------------------------------------|:------------------------------------------|
| <strong>True</strong>                        | <strong>True</strong>                         | <strong>False</strong>                    |
| <strong>True</strong>                        | <strong>False</strong>                        | <strong>True</strong>                     |
| <strong>False</strong>                       | <strong>True</strong>                         | <strong>True</strong>                     |
| <strong>False</strong>                       | <strong>False</strong>                        | <strong>False</strong>                    |

The  **Xor** operator performs as both a logical and bitwise operator. A [bit-wise comparison](vbe-glossary.md) of two [expressions](vbe-glossary.md) using exclusive-or logic to form the result, as shown in the following table:


|**If bit in  _expression1_ is**|**And bit in  _expression2_ is**|**Then  _result_ is**|
|:-----|:-----|:-----|
|0|0|0|
|0|1|1|
|1|0|1|
|1|1|0|

## Example

This example uses the  **Xor** operator to perform logical exclusion on two expressions.


```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = A > B Xor B > C    ' Returns False.
MyCheck = B > A Xor B > C    ' Returns True.
MyCheck = B > A Xor C > B    ' Returns False.
MyCheck = B > D Xor A > B    ' Returns Null.
MyCheck = A Xor B    ' Returns 2 (bitwise comparison).
```


