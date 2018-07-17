---
title: Statements or labels invalid between Select Case and first Case
keywords: vblr6.chm1011952
f1_keywords:
- vblr6.chm1011952
ms.prod: office
ms.assetid: d43e2a82-7f04-cae9-34bf-e4c819c02c74
ms.date: 06/08/2017
---


# Statements or labels invalid between Select Case and first Case

You can place nothing but a [comment](vbe-glossary.md) between the **Select Case** statement and the first **Case** clause. This error has the following cause and solution:



- You placed a statement between  **Select Case** and its first **Case** clause. For example:
    
  ```
  Select Case SomeVar 
' This is a comment and is valid. 
Stop ' Even a Stop statement is invalid here. 
Case SomeValue 
. . . 
End Select 

  ```


    The  **Select Case** statement must be immediately followed by its first **Case** statement. If the intervening[expression](vbe-glossary.md) is a comment, precede it with a comment delimiter ( **'** ). Otherwise, place the expression where it belongs or delete it.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

