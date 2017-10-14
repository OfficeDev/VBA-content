---
title: Can't rename with different drive (Error 74)
keywords: vblr6.chm1000074
f1_keywords:
- vblr6.chm1000074
ms.prod: office
ms.assetid: bba0646c-ab26-361e-5a7e-2ef6becac4a1
ms.date: 06/08/2017
---


# Can't rename with different drive (Error 74)

The  **Name** statement must rename the file to the current drive. This error has the following cause and solution:



- You tried to move a file to a different drive using the  **Name** statement. Use **FileCopy** to write the file to another drive, and then delete the old file with a **Kill** statement.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

