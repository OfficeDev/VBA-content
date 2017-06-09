---
title: Object library not registered
keywords: vblr6.chm1032797
f1_keywords:
- vblr6.chm1032797
ms.prod: office
ms.assetid: 0f2a805a-303a-43b4-6578-6c7ba3bb2627
ms.date: 06/08/2017
---


# Object library not registered

The Visual Basic for Applications [object library](vbe-glossary.md) is no longer a standalone file; it is integrated into the[dynamic-link library (DLL)](vbe-glossary.md).

In earlier versions, when you started an application that uses Visual Basic for Applications, certain object libraries were loaded. This error has the following cause and solution:




- An attempt was made to load a previous version of the Visual Basic for Applications object library (vaxxx.olb) or [host-application](vbe-glossary.md) object libraries. However, the correct language version of these object libraries could not be found in the system[registry](vbe-glossary.md).
    
    Reregister your application. On the Macintosh, delete the vba.ini file from the Macintosh Preferences folder, and restart your application.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

