---
title: "'Item' property can't be set at run time (Error 382)"
keywords: vblr6.chm382
f1_keywords:
- vblr6.chm382
ms.prod: office
ms.assetid: 20149505-5b45-6c97-228e-839bee802c62
ms.date: 06/08/2017
---


# 'Item' property can't be set at run time (Error 382)

The [property](vbe-glossary.md) is read-only at[run time](vbe-glossary.md). This error has the following cause and solution:



- You tried to set or change a property whose value can only be set at [design time](vbe-glossary.md).
    
    Remove the reference to the property from your code or change the reference to only return the value of the property at run time.
    


