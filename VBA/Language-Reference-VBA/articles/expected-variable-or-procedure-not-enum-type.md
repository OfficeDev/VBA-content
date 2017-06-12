---
title: Expected variable or procedure, not Enum type
keywords: vblr6.chm1109577
f1_keywords:
- vblr6.chm1109577
ms.prod: office
ms.assetid: e87a2297-58b5-5bf5-326c-efdeefcd9e83
ms.date: 06/08/2017
---


# Expected variable or procedure, not Enum type

The name of an  **Enum** type only appears in a statement declaring an enumeration of the type or as a qualifier. This error has the following cause and solution:



- An  **Enum** type name is used instead of the name of an enumeration variable of the type. Declare a[variable](vbe-glossary.md) of the **Enum** type or find a previous declaration in the current[scope](vbe-glossary.md) and use that variable.
    
- An  **Enum** type name is used instead of a variable or[procedure](vbe-glossary.md) name.
    
    Check the spelling of the [identifier](vbe-glossary.md) that caused the error. Use the name of a variable or procedure where you specified an **Enum** type.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

