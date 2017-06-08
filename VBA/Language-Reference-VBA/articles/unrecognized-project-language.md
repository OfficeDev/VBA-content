---
title: Unrecognized project language
keywords: vblr6.chm1032814
f1_keywords:
- vblr6.chm1032814
ms.prod: office
ms.assetid: deaf7459-f91f-2ad7-fb94-e954939a8b99
ms.date: 06/08/2017
---


# Unrecognized project language

The specified code [locale](vbe-glossary.md) for the[project](vbe-glossary.md) to be loaded isn't currently supported by this application. This error has the following causes and solutions:



- The project was created on a system that supports the code locale, but was then moved to a system where that code locale isn't recognized. For example, the ole2nls.dll on the current machine may be a version that doesn't recognize the code locale. Install the proper [dynamic-link library (DLL)](vbe-glossary.md) on the current system.
    
- The correct [object library](vbe-glossary.md) for the project was not found.
    
    Make sure the correct object libraries are available, for example, make sure your path includes their directories.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

