---
title: Invalid data format
keywords: vblr6.chm1032792
f1_keywords:
- vblr6.chm1032792
ms.prod: office
ms.assetid: 4812fe11-7137-70c3-0601-f5815827d21b
ms.date: 06/08/2017
---


# Invalid data format

The data read from a file wasn't in the expected format. This error has the following causes and solutions:



- A [project](vbe-glossary.md) file or[object library](vbe-glossary.md) file is either corrupted or in a format that can't be understood.
    
    Get a new version of the project file or object library file.
    
- You may have attempted to load an .exe file into a [module](vbe-glossary.md).
    
    Load the source code instead.
    
- You may have used the  **References** dialog box and[Object Browser](vbe-glossary.md) to add a reference to a file that isn't a valid object library or contains a Basic project in a format not supported by the[host application](vbe-glossary.md). For example, on the Windows platform, Microsoft Excel can't understand .bas or .frm files, or Microsoft Project files containing Basic code.
    
    Load the questionable file into the application in which it was created, and then save it in a compatible format. For example, object library source code can be processed through MkTypLib; and QuickBasic, and Visual Basic code can be saved in text format, and so on.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

