---
title: Can't enter break mode at this time
keywords: vblr6.chm1107949
f1_keywords:
- vblr6.chm1107949
ms.prod: office
ms.assetid: 0abba233-b7b3-8115-7575-4cde9361dc50
ms.date: 06/08/2017
---


# Can't enter break mode at this time

[Break mode](vbe-glossary.md) is the state in which a program is still running, but its activity is suspended. This error has the following cause and solution:



- You tried to enter break mode, for example, by pressing CTRL+BREAK, pressing the  **Break** button on the **Standard** toolbar or the **Debug** toolbar, or by executing a[breakpoint](vbe-glossary.md) in the running code.
    
    A change was made programmatically to the [project](vbe-glossary.md) using the extensibility (add-in) object model. This prevents the program from having execution suspended. You can continue running, or end execution, but can't suspend execution.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

