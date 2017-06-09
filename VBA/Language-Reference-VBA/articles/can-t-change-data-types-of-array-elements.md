---
title: Can't change data types of array elements
keywords: vblr6.chm1040117
f1_keywords:
- vblr6.chm1040117
ms.prod: office
ms.assetid: 909bc653-f7cc-ae95-3e43-efe06f69a629
ms.date: 06/08/2017
---


# Can't change data types of array elements

 **ReDim** can only be used to change the number of elements in an[array](vbe-glossary.md). This error has the following cause and solution:



- You tried to redeclare the [data type](vbe-glossary.md) of an array using **ReDim**.
    
    Declare a new array of the type you want, and then use the conversion functions to assign each element of the old array to the corresponding element of the new array.
    
    You can also place the array in a  **Variant** variable. This can be done with a simple assignment:
    


```vb
Dim MyVar As Variant 
MyVar = MyIntegerArray() 

  ```


    This creates a  **Variant** containing an array tagged as the type of the original array. You can then assign[variables](vbe-glossary.md) of any valid **VarType** to the elements of the array within a variant.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

