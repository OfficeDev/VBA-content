---
title: Comparison Operators
keywords: vblr6.chm1008875
f1_keywords:
- vblr6.chm1008875
ms.prod: office
ms.assetid: 9c254e88-5641-ea7d-b99a-cb614c3095a7
ms.date: 06/08/2017
---


# Comparison Operators



Used to compare [expressions](vbe-glossary.md).
 **Syntax**
 _result_**=**_expression1_ _comparisonoperator_ _expression2_
 _result_**=**_object1_**Is**_object2_
 _result_**=**_string_**Like**_pattern_
[Comparison operators](vbe-glossary.md) have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](vbe-glossary.md).|
| _expression_|Required; any expression.|
| _comparisonoperator_|Required; any comparison operator.|
| _object_|Required; any object name.|
| _string_|Required; any [string expression](vbe-glossary.md).|
| _pattern_|Required; any string expression or range of characters.|
 **Remarks**
The following table contains a list of the comparison operators and the conditions that determine whether  _result_ is **True**, **False**, or[Null](vbe-glossary.md):


|**Operator**|**True if**|**False if**|**Null if**|
|:-----|:-----|:-----|:-----|
|**&lt; (** Less than)| _expression1_ < _expression2_| _expression1_ >= _expression2_| _expression1_ or _expression2_ = **Null**|
|**&lt;= (** Less than or equal to)| _expression1_ <= _expression2_| _expression1_ > _expression2_| _expression1_ or _expression2_ = **Null**|
|**> (** Greater than)| _expression1_ > _expression2_| _expression1_ <= _expression2_| _expression1_ or _expression2_ = **Null**|
|**>= (** Greater than or equal to)| _expression1_ >= _expression2_| _expression1_ < _expression2_| _expression1_ or _expression2_ = **Null**|
|**= (** Equal to)| _expression1_ = _expression2_| _expression1_ <> _expression2_| _expression1_ or _expression2_ = **Null**|
|**&lt;> (** Not equal to)| _expression1_ <> _expression2_| _expression1_ = _expression2_| _expression1_ or _expression2_ = **Null**|

 **Note**  The  **Is** and **Like** operators have specific comparison functionality that differs from the operators in the table.

When comparing two expressions, you may not be able to easily determine whether the expressions are being compared as numbers or as strings. The following table shows how the expressions are compared or the result when either expression is not a [Variant](vbe-glossary.md):


|**If**|**Then**|
|:-----|:-----|
|Both expressions are [numeric data types](vbe-glossary.md) ([Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Date](vbe-glossary.md), [Currency](vbe-glossary.md), or [Decimal](vbe-glossary.md))|Perform a numeric comparison.|
|Both expressions are [String](vbe-glossary.md)|Perform a [string comparison](vbe-glossary.md).|
|One expression is a numeric data type and the other is a  **Variant** that is, or can be, a number|Perform a numeric comparison.|
|One expression is a numeric data type and the other is a string  **Variant** that can't be converted to a number|A  `Type Mismatch` error occurs.|
|One expression is a  **String** and the other is any **Variant** except a **Null**|Perform a string comparison.|
|One expression is [Empty](vbe-glossary.md) and the other is a numeric data type|Perform a numeric comparison, using 0 as the  **Empty** expression.|
|One expression is  **Empty** and the other is a **String**|Perform a string comparison, using a zero-length string ("") as the  **Empty** expression.|
If  _expression1_ and _expression2_ are both **Variant** expressions, their underlying type determines how they are compared. The following table shows how the expressions are compared or the result from the comparison, depending on the underlying type of the **Variant**:


|**If**|**Then**|
|:-----|:-----|
|Both  **Variant** expressions are numeric|Perform a numeric comparison.|
|Both  **Variant** expressions are strings|Perform a string comparison.|
|One  **Variant** expression is numeric and the other is a string|The numeric expression is less than the string expression.|
|One  **Variant** expression is **Empty** and the other is numeric|Perform a numeric comparison, using 0 as the  **Empty** expression.|
|One  **Variant** expression is **Empty** and the other is a string|Perform a string comparison, using a zero-length string ("") as the  **Empty** expression.|
|Both  **Variant** expressions are **Empty**|The expressions are equal.|
When a  **Single** is compared to a **Double**, the **Double** is rounded to the precision of the **Single**.
If a  **Currency** is compared with a **Single** or **Double**, the **Single** or **Double** is converted to a **Currency**. Similarly, when a **Decimal** is compared with a **Single** or **Double**, the **Single** or **Double** is converted to a **Decimal**. For **Currency**, any fractional value less than .0001 may be lost; for **Decimal**, any fractional value less than 1E-28 may be lost, or an overflow error can occur. Such fractional value loss may cause two values to compare as equal when they are not.

## Example

This example shows various uses of comparison operators, which you use to compare expressions.


```vb
Dim MyResult, Var1, Var2
MyResult = (45 < 35)    ' Returns False.
MyResult = (45 = 45)    ' Returns True.
MyResult = (4 <> 3)    ' Returns True.
MyResult = ("5" > "4")    ' Returns True.

Var1 = "5": Var2 = 4    ' Initialize variables.
MyResult = (Var1 > Var2)    ' Returns True.

Var1 = 5: Var2 = Empty
MyResult = (Var1 > Var2)    ' Returns True.

Var1 = 0: Var2 = Empty
MyResult = (Var1 = Var2)    ' Returns True.

```


