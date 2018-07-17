---
title: Array already dimensioned
keywords: vblr6.chm1040055
f1_keywords:
- vblr6.chm1040055
ms.prod: office
ms.assetid: fcf3762f-3f3f-6182-a7c9-4f055991d2c1
ms.date: 06/08/2017
---


# Array already dimensioned

A static [array](vbe-glossary.md) can only be dimensioned once. This error has the following causes and solutions:



- You attempted to change the dimensions of a static array with a  **ReDim** statement; only dynamic arrays can be redimensioned. Either remove the redimensioning or use a dynamic array. To define a dynamic array, use a **Dim**, **Public**, **Private**, or **Static** statement with empty parentheses. For example: `Dim MyArray()` In a[procedure](vbe-glossary.md), you can define a dynamic array with the  **ReDim** or **Static** statement using a[variable](vbe-glossary.md) for the number of elements:
    
```vb
Dim MyArray() 

  ```


  ```
  ReDim MyArray(n) 

  ```


     In a[procedure](vbe-glossary.md), you can define a dynamic array with the  **ReDim** or **Static** statement using a[variable](vbe-glossary.md) for the number of elements: `ReDim MyArray(n)`
    
- An  **Option Base** statement occurs after array dimensions are set. Make sure any **Option Base** statement precedes all array declarations.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

