---
title: Let Statement
keywords: vblr6.chm1008960
f1_keywords:
- vblr6.chm1008960
ms.prod: office
ms.assetid: da1ec875-3c6a-b66d-a85f-bbf33f9a307a
ms.date: 06/08/2017
---


# Let Statement

Assigns the value of an [expression](vbe-glossary.md) to a[variable](vbe-glossary.md) or[property](vbe-glossary.md).

 **Syntax**

[ **Let** ] _varname_**=**_expression_

The  **Let** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Let**|Optional. Explicit use of the  **Let**[keyword](vbe-glossary.md) is a matter of style, but it is usually omitted.|
| _varname_|Required. Name of the variable or property; follows standard variable naming conventions.|
| _expression_|Required. Value assigned to the variable or property.|
 **Remarks**
A value expression can be assigned to a variable or property only if it is of a [data type](vbe-glossary.md) that is compatible with the variable. You can't assign[string expressions](vbe-glossary.md) to numeric variables, and you can't assign[numeric expressions](vbe-glossary.md) to string variables. If you do, an error occurs at[compile time](vbe-glossary.md).
[Variant](vbe-glossary.md) variables can be assigned either string or numeric expressions. However, the reverse is not always true. Any **Variant** except a[Null](vbe-glossary.md) can be assigned to a string variable, but only a **Variant** whose value can be interpreted as a number can be assigned to a numeric variable. Use the **IsNumeric** function to determine if the **Variant** can be converted to a number.
Assigning an expression of one [numeric type](vbe-glossary.md) to a variable of a different numeric type coerces the value of the expression into the numeric type of the resulting variable.
 **Let** statements can be used to assign one record variable to another only when both variables are of the same[user-defined type](vbe-glossary.md). Use the  **LSet** statement to assign record variables of different user-defined types. Use the **Set** statement to assign object references to variables.

## Example

This example assigns the values of expressions to variables using the explicit  **Let** statement.


```vb
Dim MyStr, MyInt 
' The following variable assignments use the Let statement. 
Let MyStr = "Hello World" 
Let MyInt = 5 

```

The following are the same assignments without the  **Let** statement.




```vb
Dim MyStr, MyInt 
MyStr = "Hello World" 
MyInt = 5 

```


