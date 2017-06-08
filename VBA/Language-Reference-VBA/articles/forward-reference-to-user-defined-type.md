---
title: Forward reference to user-defined type
keywords: vblr6.chm1040069
f1_keywords:
- vblr6.chm1040069
ms.prod: office
ms.assetid: 316416e3-68d3-f5ae-88bf-6f6fa01e54a9
ms.date: 06/08/2017
---


# Forward reference to user-defined type

A [user-defined type](vbe-glossary.md) must be defined before it can be referenced. This error has the following causes and solutions:



- You declared a [variable](vbe-glossary.md) with a user-defined type before the definition of the user-defined type appears. In the following example, the variable `OtherVar` is declared before its type ( `OtherType`) is known:
    
```vb
Type MyType 
OtherVar As OtherType 
End Type 
 
Type OtherType 
WholeVar As Integer 
RealVar As Double 
End Type 

  ```


    Reposition the type definitions so that the forward reference doesn't occur.
    
- You nested a user-defined type within itself.
    
```vb
Type MyType 
MyVar As Integer 
OtherVar As MyType 
End Type 

  ```


     Remove the self-referencing nested type. This may occur indirectly if you nest a type within another type in which the first is already declared. Check the definition of each nested type to eliminate duplication.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

