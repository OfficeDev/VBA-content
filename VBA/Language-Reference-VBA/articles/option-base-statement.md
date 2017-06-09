---
title: Option Base Statement
keywords: vblr6.chm1008990
f1_keywords:
- vblr6.chm1008990
ms.prod: office
ms.assetid: 21f45e9e-2cb2-3a45-0484-d23adae77e3e
ms.date: 06/08/2017
---


# Option Base Statement

Used at [module level](vbe-glossary.md) to declare the default lower bound for[array](vbe-glossary.md) subscripts.

 **Syntax**

 **Option Base** { **0** |**1** }

 **Remarks**
Because the default base is  **0**, the **Option Base** statement is never required. If used, the[statement](vbe-glossary.md) must appear in a[module](vbe-glossary.md) before any[procedures](vbe-glossary.md).  **Option** **Base** can appear only once in a module and must precede array[declarations](vbe-glossary.md) that include dimensions.

 **Note**  The  **To** clause in the **Dim**, **Private**, **Public**, **ReDim**, and **Static** statements provides a more flexible way to control the range of an array's subscripts. However, if you don't explicitly set the lower bound with a **To** clause, you can use **Option Base** to change the default lower bound to 1. The base of an array created with the the **ParamArray** keyword is zero; **Option Base** does not affect **ParamArray** (or the **Array** function, when qualified with the name of its type library, for example **VBA.Array** ).

The  **Option Base** statement only affects the lower bound of arrays in the module where the statement is located.

## Example

This example uses the  **Option Base** statement to override the default base array subscript value of 0. The **LBound** function returns the smallest available subscript for the indicated dimension of an array. The **Option Base** statement is used at the module level only.


```vb
Option Base 1 ' Set default array subscripts to 1. 
 
Dim Lower 
Dim MyArray(20), TwoDArray(3, 4) ' Declare array variables. 
Dim ZeroArray(0 To 5) ' Override default base subscript. 
' Use LBound function to test lower bounds of arrays. 
Lower = LBound(MyArray) ' Returns 1. 
Lower = LBound(TwoDArray, 2) ' Returns 1. 
Lower = LBound(ZeroArray) ' Returns 0. 

```


