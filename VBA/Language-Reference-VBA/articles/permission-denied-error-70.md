---
title: Permission denied (Error 70)
keywords: vblr6.chm50029
f1_keywords:
- vblr6.chm50029
ms.prod: office
ms.assetid: b6822e40-c7e7-13e1-575e-632a99ad9926
ms.date: 06/08/2017
---


# Permission denied (Error 70)

An attempt was made to write to a write-protected disk or to access a locked file. This error has the following causes and solutions:



- You tried to open a write-protected file for sequential  **Output** or **Append**. Open the file for **Input** or change the write-protection attribute of the file.
    
- You tried to open a file on a disk that is write-protected for sequential  **Output** or **Append**. Remove the write-protection device from the disk or open the file for **Input**.
    
- You tried to write to a file that another process locked. Wait to open the file until the other process releases it.
    
- You attempted to access the [registry](vbe-glossary.md), but your user permissions don't include this type of registry access.
    
    On 32-bit Microsoft Windows systems, a user must have the correct permissions for access to the system registry. Change your permissions or have them changed by the system administrator.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

