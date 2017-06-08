---
title: Can't create necessary temporary file (Error 322)
keywords: vblr6.chm1000322
f1_keywords:
- vblr6.chm1000322
ms.prod: office
ms.assetid: 82464d72-90da-caea-b463-d084baf185ba
ms.date: 06/08/2017
---


# Can't create necessary temporary file (Error 322)

Creating an [executable file](vbe-glossary.md) requires creation of temporary files. This error has the following cause and solution:



- The drive that contains the directory specified by the TEMP environment variable is full. Delete files from the full drive or specify a different drive in the TEMP environment variable.
    
- The TEMP environment variable specifies an invalid or read-only drive or directory. Specify a valid drive for the TEMP environment variable or remove the read-only restriction from the currently specified drive or directory.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

