---
title: Module not found
keywords: vblr6.chm1011220
f1_keywords:
- vblr6.chm1011220
ms.prod: office
ms.assetid: bd966ba5-606c-dd48-7b2c-f27ca8e5fcee
ms.date: 06/08/2017
---


# Module not found

[Modules](vbe-glossary.md) aren't loaded from a code reference — they must be part of the[project](vbe-glossary.md). This error has the following cause and solution:



- The requested module doesn't exist in the specified project. For example, the statement  `MyModule.SomeVar = 5` generates this error when `MyModule` isn't visible in the project `MyProject`. See your [host application](vbe-glossary.md) documentation for information on including the module in the project.
    


