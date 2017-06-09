---
title: A compatible ActiveX component must be a Visual Basic executable or a DLL
keywords: vblr6.chm1015639
f1_keywords:
- vblr6.chm1015639
ms.prod: office
ms.assetid: f7057317-7bc5-4b5a-b95f-61e92a66c5f0
ms.date: 06/08/2017
---


# A compatible ActiveX component must be a Visual Basic executable or a DLL

A compatible ActiveX component is one that you specify as a compatible ActiveX component. This error has the following cause and solution:



- Visual Basic tried to access an object you specified as a compatible ActiveX component, but the file specified wasn't an [executable file](vbe-glossary.md) or[dynamic-link library (DLL)](vbe-glossary.md) created by Visual Basic.
    
    Only .exe files and DLLs created by Visual Basic are valid entries in the Compatible ActiveX Component field of the  **Project Properties** dialog box accessed through the **Project** menu. If possible, load the[project](vbe-glossary.md) into Visual Basic and choose the **Make Project.exe File** command from the **File** menu to create a Visual Basic executable file. If the file is already an executable file that wasn't created by Visual Basic, or if the file can't be loaded into Visual Basic, consult the documentation of the file to find out if it can be converted to a Visual Basic executable file or if the vendor can supply an executable file created by Visual Basic.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

