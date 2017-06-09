---
title: The edit may make the object module incompatible with the previously specified compatible ActiveX component
keywords: vblr6.chm1015638
f1_keywords:
- vblr6.chm1015638
ms.prod: office
ms.assetid: 3086d4cc-8896-e0c8-5c39-d033c2614164
ms.date: 06/08/2017
---


# The edit may make the object module incompatible with the previously specified compatible ActiveX component

If a Compatible ActiveX component already exists as a previously distributed [executable file](vbe-glossary.md) or[dynamic-link library (DLL)](vbe-glossary.md), you must be careful not to change its interface. This warning has the following cause and solution:



- You are trying to edit the code of an [object module](vbe-glossary.md) that already is represented by an executable file.
    
    If you make changes that affect the interface to the object, the class will not be upward compatible with the previous version and so it will not be possible to use the new version in place of the old version for compiled code.
    
    In Visual Basic, the name of the Compatible ActiveX component appears in the dialog box displayed when you choose  **Project Options** from the **Tools** menu.
    
     **Important**  To accept the edit, click  **OK** in the error message dialog box. If you want to undo the edit, click the **Cancel** button.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

