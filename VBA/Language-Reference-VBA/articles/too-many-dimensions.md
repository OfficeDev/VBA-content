---
title: Too many dimensions
keywords: vblr6.chm1040064
f1_keywords:
- vblr6.chm1040064
ms.prod: office
ms.assetid: c77734c3-7869-b8ee-4997-380ef68882c0
ms.date: 06/08/2017
---


# Too many dimensions

[Arrays](vbe-glossary.md) can have no more than 60 dimensions. This error has the following causes and solutions:



- You tried to declare an array with more than 60 dimensions. Reduce the number of dimensions.
    
- Your array declaration is within the specified limits, but there isn't enough memory to actually create the array. Either make more memory available or reduce the number of dimensions. If your array is an array of  **Variant** type or an array contained within a **Variant**, you may be able to create the array with the same number of dimensions by redeclaring it with the[data type](vbe-glossary.md) of its elements. For example, if it contains only integers, declaring it as an array of **Integer** type uses less memory than if each element is a **Variant**.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

