---
title: Can't perform requested operation (Error 17)
keywords: vblr6.chm1011338
f1_keywords:
- vblr6.chm1011338
ms.prod: office
ms.assetid: 4cde1fa7-b509-4d69-2157-7fb0a429d99f
ms.date: 06/08/2017
---


# Can't perform requested operation (Error 17)

An operation can't be carried out if it would invalidate the current state of the [project](vbe-glossary.md). This error has the following cause and solution:



- The requested operation would invalidate the current state of the project. For example, the error occurs if you use the  **References** dialog box to add a reference to a new project or[object library](vbe-glossary.md) while a program is in[break mode](vbe-glossary.md).
    
    Stop execution of the current code, and then retry the operation.
    
- An attempt was made to programmatically modify currently running code. For example, your code may have tried to read code from a disk file into a currently running [module](vbe-glossary.md).
    
    Although you can modify modules in the project while they aren't actually running, you can't make modifications to a running module. To make such changes, you must stop the module from running, make the additions or changes, and then restart execution.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

