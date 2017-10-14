---
title: Using Data Types Efficiently
keywords: vbcn6.chm1009792
f1_keywords:
- vbcn6.chm1009792
ms.prod: office
ms.assetid: 8777099a-623e-4fce-ef3b-beecf10cb603
ms.date: 06/08/2017
---


# Using Data Types Efficiently

Unless otherwise specified, undeclared [variables](vbe-glossary.md) are assigned the[Variant data type](vbe-glossary.md). This data type makes it easy to write programs, but it is not always the most efficient data type to use.

You should consider using other data types if:




- Your program is very large and uses many variables.
    
- Your program must run as quickly as possible.
    
- You write data directly to random-access files.
    

In addition to  **Variant**, supported data types include **Byte**, **Boolean**, **Integer**, **Long**, **Single**, **Double**, **Currency**, **Decimal**, **Date**, **Object**, and **String**. Use the **Dim** statement to declare a variable of a specific type, for example:



```vb
Dim X As Integer 

```

This statement declares that a variable  `X` is an integer — a whole number between -32,768 and 32,767. If you try to set `X` to a number outside that range, an error occurs. If you try to set `X` to a fraction, the number is rounded. For example:



```
X = 32768        ' Causes error. 
X = 5.9        ' Sets x to 6. 

```


