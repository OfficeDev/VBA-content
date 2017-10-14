---
title: Variable not yet created in this context
keywords: vblr6.chm1040363
f1_keywords:
- vblr6.chm1040363
ms.prod: office
ms.assetid: 93dc4805-7ce4-0240-7bc7-e5bc593dfbf5
ms.date: 06/08/2017
---


# Variable not yet created in this context

A [variable](vbe-glossary.md) has to be created before it can be displayed in the **Watch** window or the **Immediate** window, and before it can have values assigned to it in the **Immediate** window. This error has the following causes and solutions:



- You tried to display the value of a local variable that you just entered in your code before executing at least a  **Single Step** command in[break mode](vbe-glossary.md).
    
    Step into the code to force compilation of the new statement.
    
- You tried to display the value of a local variable that you just added in a [procedure](vbe-glossary.md) farther down the call chain by moving to the procedure using the **Calls** dialog box.
    
    You have to actually return to the procedure before you can display the variable in its context.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

