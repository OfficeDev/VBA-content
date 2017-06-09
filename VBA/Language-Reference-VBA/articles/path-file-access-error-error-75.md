---
title: Path/File access error (Error 75)
keywords: vblr6.chm1011249
f1_keywords:
- vblr6.chm1011249
ms.prod: office
ms.assetid: 5c1e151a-facd-6e55-d075-f7faef4a2793
ms.date: 06/08/2017
---


# Path/File access error (Error 75)

During a file-access or disk-access operation, for example,  **Open**, **MkDir**, **ChDir**, or **RmDir**, the operating system couldn't make a connection between the path and the file name. This error has the following causes and solutions:



- The file specification isn't correctly formatted. A file name can contain a fully qualified (absolute) or relative path. A fully qualified path starts with the drive name (if the path is on another drive) and lists the explicit path from the root to the file. Any path that isn't fully qualified is relative to the current drive and directory.
    
- You attempted to save a file that would replace an existing read-only file. Change the read-only attribute of the target file or save the file with a different file name.
    
- You attempted to open a read-only file in sequential  **Output** or **Append** mode. Open the file in **Input** mode or change the read-only attribute of the file.
    
- You attempted to change a Visual Basic project within a database or document. You can't make design changes to the project.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

