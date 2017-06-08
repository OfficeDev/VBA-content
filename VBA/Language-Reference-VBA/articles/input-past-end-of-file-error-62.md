---
title: Input past end of file (Error 62)
keywords: vblr6.chm50021
f1_keywords:
- vblr6.chm50021
ms.prod: office
ms.assetid: cd2a6984-2dae-66f0-ee55-14372a1d5f0a
ms.date: 06/08/2017
---


# Input past end of file (Error 62)

You can't read past the end-of-file position. This error has the following cause and solution:



- An  **Input #** or **Line Input #** statement is reading from a file in which all data has been read or from an empty file. Use the **EOF** function immediately before the **Input #** statement to detect the end of file.
    
- You used the  **EOF** function with a file opened for **Binary** access. **EOF** only works with files opened for sequential **Input** access. Use **Seek** and **Loc** with files opened for **Binary** access.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

