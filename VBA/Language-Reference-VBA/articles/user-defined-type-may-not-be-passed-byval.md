---
title: User-defined type may not be passed ByVal
keywords: vblr6.chm1040140
f1_keywords:
- vblr6.chm1040140
ms.prod: office
ms.assetid: 1fbfeef6-b92d-03ca-aeec-4cf4c0d8d972
ms.date: 06/08/2017
---


# User-defined type may not be passed ByVal

[User-defined types](vbe-glossary.md) can only be passed[by reference](vbe-glossary.md) (the default), not[by value](vbe-glossary.md). The error may not be reported until the call is made. This error has the following cause and solution:



- You placed a  **ByVal** keyword in the definition of a[parameter](vbe-glossary.md) that represented a user-defined type.
    
    Remove the  **ByVal** keyword. To keep changes from being propagated back to the caller, **Dim** a temporary[variable](vbe-glossary.md) of the type and pass the temporary variable into the[procedure](vbe-glossary.md).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

