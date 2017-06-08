---
title: The .VBP file for this project contains an invalid or corrupt library references ID
keywords: vblr6.chm1035011
f1_keywords:
- vblr6.chm1035011
ms.prod: office
ms.assetid: 0509d8f4-deae-f460-a376-11c637cc6ece
ms.date: 06/08/2017
---


# The .VBP file for this project contains an invalid or corrupt library references ID

When you save a [project](vbe-glossary.md) for which a reference has been selected from the **References** dialog box, an entry is made in the project's .vbp file (called the .mak file in earlier versions of Visual Basic). For example, the entry for a data access object is:


```text
Reference=*\G{00025E01-0000-0000-C000-000000000046}#0.0#0#C:\WINDOWS\SYSTEM\DAO2516.DLL#Microsoft 
DAO 2.5 Object Library 

```


This error occurs when such a reference has been edited or corrupted. This error has the following cause and solution:



- A reference in the .vbp file has become invalid. Delete the incorrect line from the .vbp file and check the appropriate [object library](vbe-glossary.md) in the **References** dialog box from the **Tools** menu. Then save the project, and the correct information will be entered in the .vbp file.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

