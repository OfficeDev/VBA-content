---
title: + Operator
keywords: vblr6.chm1008845
f1_keywords:
- vblr6.chm1008845
ms.prod: office
ms.assetid: fa6ce797-a49c-af99-4ab5-112056c2a584
ms.date: 06/08/2017
---


# + Operator



Used to sum two numbers.
 **Syntax**
 _result_ **=** _expression1 + expression2_.
The  **+** operator syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                       |
|:----------------------|:---------------------------------------------------|
| <em>result</em>       | Required; any numeric [variable](vbe-glossary.md). |
| <em>expression1</em>  | Required; any [expression](vbe-glossary.md).       |
| <em>expression2</em>  | Required; any expression.                          |

 **Remarks**
When you use the  **+** operator, you may not be able to determine whether addition or string concatenation will occur. Use the **&;** operator for concatenation to eliminate ambiguity and provide self-documenting code.
If at least one expression is not a [Variant](vbe-glossary.md), the following rules apply:


| <strong>If</strong>                                                                                                                                                                                                                                                                                              | <strong>Then</strong>                                          |
|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|:---------------------------------------------------------------|
| Both expressions are [numeric data types](vbe-glossary.md) ([Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Date](vbe-glossary.md), [Currency](vbe-glossary.md), or [Decimal](vbe-glossary.md)) | Add.                                                           |
| Both expressions are [String](vbe-glossary.md)                                                                                                                                                                                                                                                                   | Concatenate.                                                   |
| One expression is a numeric data type and the other is any  <strong>Variant</strong> except [Null](vbe-glossary.md)                                                                                                                                                                                              | Add.                                                           |
| One expression is a  <strong>String</strong> and the other is any <strong>Variant</strong> except <strong>Null</strong>                                                                                                                                                                                          | Concatenate.                                                   |
| One expression is an [Empty](vbe-glossary.md) <strong>Variant</strong>                                                                                                                                                                                                                                           | Return the remaining expression unchanged as  <em>result</em>. |
| One expression is a numeric data type and the other is a  <strong>String</strong>                                                                                                                                                                                                                                | A  `Type mismatch`error occurs.                                |
| Either expression is  <strong>Null</strong>                                                                                                                                                                                                                                                                      | <em>result</em> is <strong>Null</strong>.                      |

If both expressions are  **Variant** expressions, the following rules apply:


| <strong>If</strong>                                                           | <strong>Then</strong> |
|:------------------------------------------------------------------------------|:----------------------|
| Both  <strong>Variant</strong> expressions are numeric                        | Add.                  |
| Both  <strong>Variant</strong> expressions are strings                        | Concatenate.          |
| One  <strong>Variant</strong> expression is numeric and the other is a string | Add.                  |

For simple arithmetic addition involving only expressions of numeric data types, the [data type](vbe-glossary.md) of _result_ is usually the same as that of the most precise expression. The order of precision, from least to most precise, is **Byte**, **Integer**, **Long**, **Single**, **Double**, **Currency**, and **Decimal**. The following are exceptions to this order:


| <strong>If</strong>                                                                                                                                     | <strong>Then  <em>result</em> is</strong>          |
|:--------------------------------------------------------------------------------------------------------------------------------------------------------|:---------------------------------------------------|
| A  <strong>Single</strong> and a <strong>Long</strong> are added,                                                                                       | a  <strong>Double</strong>.                        |
| The data type of  <em>result</em> is a <strong>Long</strong>, <strong>Single</strong>, or <strong>Date</strong> variant that overflows its legal range, | converted to a  <strong>Double</strong> variant.   |
| The data type of  <em>result</em> is a <strong>Byte</strong> variant that overflows its legal range,                                                    | converted to an  <strong>Integer</strong> variant. |
| The data type of  <em>result</em> is an <strong>Integer</strong> variant that overflows its legal range,                                                | converted to a  <strong>Long</strong> variant.     |
| A  <strong>Date</strong> is added to any data type,                                                                                                     | a  <strong>Date</strong>.                          |

If one or both expressions are  **Null** expressions, _result_ is **Null**. If both expressions are **Empty**, _result_ is an **Integer**. However, if only one expression is **Empty**, the other expression is returned unchanged as _result_.

 **Note**  The order of precision used by addition and subtraction is not the same as the order of precision used by multiplication.


## Example

This example uses the  **+** operator to sum numbers. The **+** operator can also be used to concatenate strings. However, to eliminate ambiguity, you should use the **&;** operator instead. If the components of an expression created with the **+** operator include both strings and numerics, the arithmetic result is assigned. If the components are exclusively strings, the strings are concatenated.


```vb
Dim MyNumber, Var1, Var2
MyNumber = 2 + 2    ' Returns 4.
MyNumber = 4257.04 + 98112    ' Returns 102369.04.

Var1 = "34": Var2 = 6    ' Initialize mixed variables.
MyNumber = Var1 + Var2    ' Returns 40.

Var1 = "34": Var2 = "6"    ' Initialize variables with strings.
MyNumber = Var1 + Var2    ' Returns "346" (string concatenation).
```


