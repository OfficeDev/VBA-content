---
title: Specified DLL function not found (Error 453)
keywords: vblr6.chm1011197
f1_keywords:
- vblr6.chm1011197
ms.prod: office
ms.assetid: 5065c4a8-e5fb-2c47-0c8b-25afcbe2c2f3
ms.date: 06/08/2017
---


# Specified DLL function not found (Error 453)

The [dynamic-link library (DLL)](vbe-glossary.md) in a user library reference was found, but the DLL function specified wasn't found within the DLL. This error has the following causes and solutions:



- You specified an invalid ordinal in the function declaration. Check for the proper ordinal or call the function by name.
    
- You gave the right DLL name, but it isn't the version that contains the specified function. You may have the correct version on your machine, but if the directory containing the wrong version precedes the directory containing the correct one in your path, the wrong DLL is accessed. Check your machine for different versions. If you have an early version, contact the supplier for a later version.
    
- If you are working on a 32-bit Microsoft Windows platform, both the DLL name and alias (if used) must be correct. Make sure the DLL name and alias are correct.
    
- Some 32-bit DLLs contain functions with slightly different versions to accommodate both [Unicode](vbe-glossary.md) and[ANSI](vbe-glossary.md) strings. An "A" at the end of the function name specifies the ANSI version. A "W" at the end of the function name specifies the Unicode version.
    
    If the function takes string-type arguments, try appending an "A" to the function name.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

