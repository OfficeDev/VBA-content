---
title: Declaring Constants
keywords: vbcn6.chm1076698
f1_keywords:
- vbcn6.chm1076698
ms.prod: office
ms.assetid: c1b65bc4-1e94-828c-67bf-357a75261657
ms.date: 06/08/2017
---


# Declaring Constants

By declaring a [constant](vbe-glossary.md), you can assign a meaningful name to a value. You use the  **Const** statement to declare a constant and set its value. After a constant is declared, it cannot be modified or assigned a new value.

You can declare a constant within a [procedure](vbe-glossary.md) or at the top of a [module](vbe-glossary.md), in the Declarations section. [Module-level](vbe-glossary.md) constants are private by default. To declare a public module-level constant, precede the **Const** statement with the **Public** [keyword](vbe-glossary.md). You can explicitly declare a private constant by preceding the  **Const** statement with the **Private** keyword to make it easier to read and interpret your code. For more information, see "Understanding Scope and Visibility" in Visual Basic Help.

The following example declares the  **Public** constant `conAge` as an **Integer** and assigns it the value as an **Integer** and assigns it the value `34`.




```vb
Public Const conAge As Integer = 34
```

Constants can be declared as one of the following data types:  **Boolean**, **Byte**, **Integer**, **Long**, **Currency**, **Single**, **Double**, **Date**, **String**, or **Variant**. Because you already know the value of a constant, you can specify the data type in a **Const** statement. For more information on data types, see "Data Type Summary" in Visual Basic Help.
You can declare several constants in one statement. To specify a data type, you must include the data type for each constant. In the following statement, the constants  `conAge` and and `conWage` are declared as **Integer**.



```vb
Const conAge As Integer = 34, conWage As Currency = 35000
```


