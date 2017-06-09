---
title: Object library's language setting incompatible with current project
keywords: vblr6.chm1040340
f1_keywords:
- vblr6.chm1040340
ms.prod: office
ms.assetid: 4427c8c1-9990-0761-5f5b-2c58ba6eb329
ms.date: 06/08/2017
---


# Object library's language setting incompatible with current project

The reference couldn't be added. This error has the following cause and solution:



- You attempted to add a reference to an [object library](vbe-glossary.md) whose[locale](vbe-glossary.md) isn't compatible with the locale of the current[project](vbe-glossary.md). The reference was not added. To use that object library, a project whose locale is compatible with it must be created.
    
    Try registering both Visual Basic for Applications and the [host application](vbe-glossary.md) for the given language. The object library then becomes available in the **References** dialog box.
    
     **Note**  When Visual Basic is the host application, it isn't possible to change a project's language setting. Any object libraries used must be compatible with the English/U.S. setting.


