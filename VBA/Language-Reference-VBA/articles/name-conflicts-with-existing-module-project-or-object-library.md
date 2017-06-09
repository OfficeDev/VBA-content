---
title: Name conflicts with existing module, project, or object library
keywords: vblr6.chm1032813
f1_keywords:
- vblr6.chm1032813
ms.prod: office
ms.assetid: 0096e260-4af8-e133-1d64-6e606f371df2
ms.date: 06/08/2017
---


# Name conflicts with existing module, project, or object library

[Modules](vbe-glossary.md), [object libraries](vbe-glossary.md), and [referenced projects](vbe-glossary.md) must be uniquely named within a[project](vbe-glossary.md). This error has the following causes and solutions:



- There is already a module, project, or object library with this name referenced in this project. A file name extension isn't considered part of the name, so different extensions can't be used to distinguish one file from another. Use a different name for one of the duplicate module, project, or object library references.
    
- You attempted to add a reference to a project or object library whose file name (without an extension) is the same as the name of one of the current project's modules. Change either the module name or the name of the file that could not be added.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

