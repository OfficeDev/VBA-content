---
title: Procedure type mismatch
keywords: vblr6.chm1035020
f1_keywords:
- vblr6.chm1035020
ms.prod: office
ms.assetid: 6190c447-7ff5-67b1-25ec-19e1b9628238
ms.date: 06/08/2017
---


# Procedure type mismatch

You are using one kind of  **Property** procedure where a different kind is expected. This error has the following causes and solutions:



- You are trying to write to a [property](vbe-glossary.md) that is read-only.
    
    If the only [property procedure](vbe-glossary.md) defined for the property is a **Property Get**, you can't assign a value to the property. Either write an appropriate **Property Let** procedure, or don't attempt to write to the property.
    
- You are trying to read a property that is write-only. If the only property procedure defined for the property is a  **Property Let**, you can't read the value of the property. Either write an appropriate **Property Get** procedure, or don't attempt to write to the property.
    
- You are trying to set a reference but the property has only  **Property Get** or **Property Let** procedures. Either write a **Property Set** procedure for the property, or don't try to set a reference to it.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

