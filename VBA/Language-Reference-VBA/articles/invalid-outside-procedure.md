---
title: Invalid outside procedure
keywords: vblr6.chm1040051
f1_keywords:
- vblr6.chm1040051
ms.prod: office
ms.assetid: 46c00b2b-c656-9ad4-bff9-d341a6a7ecd5
ms.date: 06/08/2017
---


# Invalid outside procedure

The statement must occur within a  **Sub** or **Function**, or a property procedure ( **Property Get**, **Property Let**, **Property Set** ). This error has the following cause and solution:



- An executable statement,  **Static** or **ReDim**, appears at[module level](vbe-glossary.md).
    
     **Static** is unnecessary at module level, since all module-level[variables](vbe-glossary.md) are static. Use **Dim** instead of **ReDim** at module level. To create a dynamic[array](vbe-glossary.md) at module level, declare it with **Dim** using empty parentheses.
    
     **Note**  At module level, you can use only [comments](vbe-glossary.md) and declarative statements, such as **Const**, **Declare**, **Def**_type_, **Dim**, **Option Base**, **Option Compare**, **Option Explicit**, **Option Private**, **Private**, **Public**, and **Type**. The **Sub**, **Function**, and **Property** statements occur outside the body of their[procedures](vbe-glossary.md), but within the procedure declaration.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

