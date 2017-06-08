---
title: "Seek failed: can't read/write from the disk"
keywords: vblr6.chm1011272
f1_keywords:
- vblr6.chm1011272
ms.prod: office
ms.assetid: b91ba9ad-672d-a2a8-ffa2-4f19cdf2119e
ms.date: 06/08/2017
---


# Seek failed: can't read/write from the disk

 **Seek** statements are carried out directly to disk. This error has the following causes and solutions:



- You attempted to read from a disk or file that is write-protected, read-only, or locked. Remove the write-protected attribute or change the read-only attribute or lock. Note that if the file is locked by another process, you can't remove the lock.
    
- The file has become unavailable, for example, if a removable disk has been physically changed. If the file has been moved to another disk, access it from there. Otherwise, you can't access the file.
    
- You attempted to read from a [project](vbe-glossary.md) file, an[object library](vbe-glossary.md), or a [type library](vbe-glossary.md), but the file has been corrupted.
    
    Obtain a new copy of the library or project file.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

