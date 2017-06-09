---
title: Breakpoint not allowed on this line
keywords: vblr6.chm1011228
f1_keywords:
- vblr6.chm1011228
ms.prod: office
ms.assetid: fee3a55e-9598-3c71-f855-8f272cb19d96
ms.date: 06/08/2017
---


# Breakpoint not allowed on this line

[Breakpoints](vbe-glossary.md) can only be placed on certain parts of statements. This error has the following causes:



- You tried to place a breakpoint on a line that can't accept a breakpoint, for example:
    
    
    
      - A line that contains only [comments](vbe-glossary.md).
    
  - A line that contains only [line labels](vbe-glossary.md).
    
  - A line that contains only [declarations](vbe-glossary.md) ( **Const**, **Dim**, **Static**, **Type**, and so on).
    
  - Any line in a hidden [module](vbe-glossary.md).
    
  - Any line in the  **Immediate** window.
    

    
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

