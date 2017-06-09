---
title: Statement invalid outside Type block
keywords: vblr6.chm1040053
f1_keywords:
- vblr6.chm1040053
ms.prod: office
ms.assetid: 287d4cf7-257a-2cc4-2e5d-42e578c8b862
ms.date: 06/08/2017
---


# Statement invalid outside Type block

The syntax for declaring [variables](vbe-glossary.md) outside a **Type...End Type** statement block is different from the syntax for declaring the elements of the[user-defined type](vbe-glossary.md). This error has the following causes and solutions:



- You tried to declare a variable outside a  **Type...End Type** block or outside a statement. When declaring a variable with an **As** clause outside a **Type...End Type** block, use one of the declaration statements, **Dim**, **ReDim**, **Static**, **Public**, or **Private**. For example, the first declaration of `MyVar` in the following code generates this error; the second and third declarations of `MyVar` are valid:
    
```vb
MyVar As Double ' Invalid declaration syntax. 
 
Dim MyVar As Double 
 
Type AType 
MyVar As Double ' This is valid declaration syntax 
End Type ' because it's inside a Type block. 

  ```


    
    
- You used an  **End Type** statement without a corresponding **Type** statement. Check for an unmatched **End Type**, and either precede its block with a **Type** statement, or delete the **End Type** statement if it isn't needed.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

