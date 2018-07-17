---
title: Understanding Variants
keywords: vbcn6.chm1076678
f1_keywords:
- vbcn6.chm1076678
ms.prod: office
ms.assetid: 0f8d3917-0ca3-0a67-2c3d-48883f4a24f1
ms.date: 06/08/2017
---


# Understanding Variants

The  **Variant** data type is automatically specified if you don't specify a [data type](vbe-glossary.md) when you declare a [constant](vbe-glossary.md), [variable](vbe-glossary.md), or [argument](vbe-glossary.md). Variables declared as the  **Variant** data type can contain string, date, time, Boolean, or numeric values, and can convert the values they contain automatically. Numeric **Variant** values require 16 bytes of memory (which is significant only in large [procedures](vbe-glossary.md) or complex [modules](vbe-glossary.md)) and they are slower to access than explicitly typed variables of any other type. You rarely use the  **Variant** data type for a constant. String **Variant** values require 22 bytes of memory.

The following statements create  **Variant** variables:



```vb
Dim myVar 
Dim yourVar As Variant 
theVar = "This is some text." 

```

The last statement does not explicitly declare the variable , but rather declares the variable implicitly, or automatically. Variables that are declared implicitly are specified as the  **Variant** data type.

 **Tip**  If you specify a data type for a variable or argument, and then use the wrong data type, a data type error will occur. To avoid data type errors, either use only implicit variables (the  **Variant** data type) or explicitly declare all your variables and specify a data type. The latter method is preferred.


