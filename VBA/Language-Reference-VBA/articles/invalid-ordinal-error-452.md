---
title: Invalid ordinal (Error 452)
keywords: vblr6.chm1000452
f1_keywords:
- vblr6.chm1000452
ms.prod: office
ms.assetid: 10f033c8-d76e-710d-4014-ba2d171745a9
ms.date: 06/08/2017
---


# Invalid ordinal (Error 452)

Your call to a [dynamic-link library (DLL)](vbe-glossary.md) indicated to use a number instead of a procedure name, using the **#**_num_ syntax. This error has the following causes and solutions:



- An attempt to convert the  _num_ expression to an ordinal failed. Make sure the[expression](vbe-glossary.md) represents a valid number or call the[procedure](vbe-glossary.md) by name.
    
- The  _num_ specified doesn't specify any function in the DLL. Make sure _num_ identifies a valid function in the DLL.
    
- A [type library](vbe-glossary.md) has an invalid declaration resulting in internal use of an invalid ordinal number.
    
    [Comment](vbe-glossary.md) out code to isolate the procedure call causing the problem. Write a **Declare** statement for the procedure and report the problem to the type library vendor.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

