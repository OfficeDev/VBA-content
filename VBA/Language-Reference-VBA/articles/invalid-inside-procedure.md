---
title: Invalid inside procedure
keywords: vblr6.chm1011202
f1_keywords:
- vblr6.chm1011202
ms.prod: office
ms.assetid: ba314d7c-1d01-6b99-f80b-b1c18b1bef32
ms.date: 06/08/2017
---


# Invalid inside procedure

The statement can't occur in a  **Sub** or **Function** procedure. This error has the following cause and solution:



- One of the following statements appears in a [procedure](vbe-glossary.md):  **Declare**, **Def**_type_, **Private**, **Public**, **Option Base**, **Option Compare**, **Option Explicit**, **Option Private**, **Enum** and **Type**.
    
    Remove the statement from the procedure. The statements can be placed at [module level](vbe-glossary.md).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

