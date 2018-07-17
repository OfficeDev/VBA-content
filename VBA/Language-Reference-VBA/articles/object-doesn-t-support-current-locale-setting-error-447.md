---
title: Object doesn't support current locale setting (Error 447)
keywords: vblr6.chm1011333
f1_keywords:
- vblr6.chm1011333
ms.prod: office
ms.assetid: 5039df77-9505-ff20-3823-875bc2701cde
ms.date: 06/08/2017
---


# Object doesn't support current locale setting (Error 447)

Not all objects support all [locale](vbe-glossary.md) settings. This error has the following causes and solutions:



- You attempted to access an object that doesn't support the locale setting for the current [project](vbe-glossary.md). For example, if your current project has the locale setting Canadian French, the object you are trying to access must support that locale setting.
    
    Check which locale settings the object supports.
    
- The object relies on national language support in a [dynamic-link library (DLL)](vbe-glossary.md), for example, OLE2NLS.DLL, that may be out of date.
    
    Obtain a more recent version of the DLL, one that supports the current project locale.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

