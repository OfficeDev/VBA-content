---
title: Invalid picture type (Error 485)
keywords: vblr6.chm1117801
f1_keywords:
- vblr6.chm1117801
ms.prod: office
ms.assetid: 3b0c25f3-8faa-efe4-1a77-676696dca3d1
ms.date: 06/08/2017
---


# Invalid picture type (Error 485)

The resource file picture format you tried to load doesn't match the specified property of the object. This error has the following causes and solutions:



- You tried to use the  **LoadResPicture** method to load a bitmap resource as the **Icon** property of a form. Change the property to the **Picture** property or change the _format_[argument](vbe-glossary.md) of **LoadResPicture** to **vbResIcon**.
    
- You tried to use the  **LoadResPicture** method to load a cursor resource as some property of an object or control other than the **MousePointer** property. Change the property reference to **MousePointer**.
    


