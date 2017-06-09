---
title: Current module does not support Print method
keywords: vblr6.chm1011234
f1_keywords:
- vblr6.chm1011234
ms.prod: office
ms.assetid: 30f14bb8-ebc6-cbd7-e1f2-e557836c93a9
ms.date: 06/08/2017
---


# Current module does not support Print method

Not all [methods](vbe-glossary.md) and[properties](vbe-glossary.md) are appropriate in all[modules](vbe-glossary.md). This error has the following causes and solutions:



- You tried to use the  **Print** method on an object that can't display anything. For example, you can't use the **Print** method without qualification in a[standard module](vbe-glossary.md).
    
    Remove the reference to the  **Print** method, or qualify it with an appropriate object. For example, qualify it with the **Debug** object to display its arguments in the **Immediate** window during debugging.
    
- You tried to use the  **Line**, **Circle**, **PSet**, or **Scale** method on an object that can't accept them. For example, they can't appear unqualified in a standard module or an Automation[class module](vbe-glossary.md).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

