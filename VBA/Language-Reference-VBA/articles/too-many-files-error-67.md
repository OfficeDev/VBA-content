---
title: Too many files (Error 67)
keywords: vblr6.chm1011283
f1_keywords:
- vblr6.chm1011283
ms.prod: office
ms.assetid: d1ef7ab6-a99d-02ab-61ac-1743b95897f2
ms.date: 06/08/2017
---


# Too many files (Error 67)

There is a limit to the number of disk files that can be open at one time. This error has the following causes and solutions:



- MS-DOS operating system: More files have been created in the root directory than the operating system permits. The MS-DOS operating system limits the number of files that can be in the root directory, usually 512. If your program is opening, closing, or saving files in the root directory, change your program so that it uses a subdirectory.
    
- MS-DOS operating system: More files have been opened than the number specified in the  **files=** setting in your CONFIG.SYS file. Increase the number specified in the **files=** setting in your CONFIG.SYS file and restart your computer.
    
- Macintosh: Your program tried to open more than 40 files. On the Macintosh, the standard limit is 40 files. This limit can be changed using a utility to modify the  **MaxFiles** parameter of the boot block.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

