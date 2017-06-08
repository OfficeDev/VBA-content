---
title: Can't save file to TEMP directory (Error 735)
keywords: vblr6.chm50052
f1_keywords:
- vblr6.chm50052
ms.prod: office
ms.assetid: 587d741e-c2ad-e5c7-5390-dadc1bea4acb
ms.date: 06/08/2017
---


# Can't save file to TEMP directory (Error 735)

Components often need to save temporary information to disk. This error has the following cause and solution:



- Component can't find a directory named TEMP. Create a directory named TEMP and set the TEMP environment variable equal to its path.
    
- The drive or partition containing the TEMP directory lacks sufficient space to save information. Make some space on the drive by erasing unnecessary files, or create a TEMP directory on another partition and set the TEMP environment variable equal to its path.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

