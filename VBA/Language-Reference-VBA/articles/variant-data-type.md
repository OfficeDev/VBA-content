---
title: Variant Data Type
keywords: vblr6.chm1009056
f1_keywords:
- vblr6.chm1009056
ms.prod: office
ms.assetid: 19750b07-c2bf-dff7-67a1-91b06338cbc6
ms.date: 06/08/2017
---


# Variant Data Type

The **Variant** data type is the [data type](vbe-glossary.md) for all [variables](vbe-glossary.md) that are not explicitly declared as some other type (using [statements](vbe-glossary.md) such as **Dim**, **Private**, **Public**, or **Static** ). The **Variant** data type has no [type-declaration character](vbe-glossary.md).

A **Variant** is a special data type that can contain any kind of data except fixed-length [String](vbe-glossary.md) data. ( **Variant** types now support [user-defined types](vbe-glossary.md).) A **Variant** can also contain the special values [Empty](vbe-glossary.md), **Error**, **Nothing**, and [Null](vbe-glossary.md). You can determine how the data in a **Variant** is treated using the **VarType** function or **TypeName** function.

Numeric data can be any integer or real number value ranging from -1.797693134862315E308 to -4.94066E-324 for negative values and from 4.94066E-324 to 1.797693134862315E308 for positive values. Generally, numeric **Variant** data is maintained in its original data type within the **Variant**. For example, if you assign an [Integer](vbe-glossary.md) to a **Variant**, subsequent operations treat the **Variant** as an **Integer**. However, if an arithmetic operation is performed on a **Variant** containing a [Byte](vbe-glossary.md), an **Integer**, a [Long](vbe-glossary.md), or a [Single](vbe-glossary.md), and the result exceeds the normal range for the original data type, the result is promoted within the **Variant** to the next larger data type. A **Byte** is promoted to an **Integer**, an **Integer** is promoted to a **Long**, and a **Long** and a **Single** are promoted to a [Double](vbe-glossary.md). An error occurs when **Variant** variables containing [Currency](vbe-glossary.md), [Decimal](vbe-glossary.md), and **Double** values exceed their respective ranges.

You can use the **Variant** data type in place of any data type to work with data in a more flexible way. If the contents of a **Variant** variable are digits, they may be either the string representation of the digits or their actual value, depending on the context. For example:



```vb
Dim MyVar As Variant 
MyVar = 98052 

```

In the preceding example,  `MyVar` contains a numeric representation — the actual value `98052`. Arithmetic operators work as expected on  **Variant** variables that contain numeric values or string data that can be interpreted as numbers. If you use the **+** operator to add contains a numeric representation — the actual value `98052`. Arithmetic operators work as expected on  **Variant** variables that contain numeric values or string data that can be interpreted as numbers. If you use the **+** operator to add `MyVar` to another **Variant** containing a number or to a variable of a [numeric type](vbe-glossary.md), the result is an arithmetic sum.
The value [Empty](vbe-glossary.md) denotes a **Variant** variable that hasn't been initialized (assigned an initial value). A **Variant** containing **Empty** is 0 if it is used in a numeric context and a zero-length string ("") if it is used in a string context.
Don't confuse  **Empty** with [Null](vbe-glossary.md).  **Null** indicates that the **Variant** variable intentionally contains no valid data.
In a  **Variant**, **Error** is a special value used to indicate that an error condition has occurred in a[procedure](vbe-glossary.md). However, unlike for other kinds of errors, normal application-level error handling does not occur. This allows you, or the application itself, to take some alternative action based on the error value.  **Error** values are created by converting real numbers to error values using the **CVErr** function.



