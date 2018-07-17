---
title: This component doesn't raise any events
keywords: vblr6.chm1109576
f1_keywords:
- vblr6.chm1109576
ms.prod: office
ms.assetid: ab95a71c-b368-ed4b-de0c-06a2fb41382f
ms.date: 06/08/2017
---


# This component doesn't raise any events

An event [procedure](vbe-glossary.md) must correspond to an event that can be raised by an object. This error has the following cause and solution:



- You wrote an event procedure for an object that doesn't raise events. You can't write an event procedure that doesn't correspond to an event.
    
- You tried to use  **WithEvents** on a[class](vbe-glossary.md) that doesn't raise events.
    
    You can't use  **WithEvents** on a class that doesn't raise events.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

