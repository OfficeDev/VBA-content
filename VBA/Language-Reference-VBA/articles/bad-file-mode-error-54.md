---
title: Bad file mode (Error 54)
keywords: vblr6.chm1011083
f1_keywords:
- vblr6.chm1011083
ms.prod: office
ms.assetid: cc5a69ce-9d99-0f20-ac36-9a6e512ec032
ms.date: 06/08/2017
---


# Bad file mode (Error 54)

Statements used in manipulating file contents must be appropriate to the mode in which the file was opened. This error has the following causes and solutions:



- A  **Put** or **Get** statement is specifying a sequential file. **Put** and **Get** can only refer to files opened for **Random** or **Binary** access.
    
- A  **Print #** statement specifies a file opened for an access mode other than **Output** or **Append**. Use a different statement to place data in the file or reopen the file in an appropriate mode.
    
- An  **Input #** statement specifies a file opened for an access mode other than **Input**. Use a different statement to place data in the file or reopen the file in **Input** mode.
    
- You attempted to write to a read-only file. Change the read/write status of the file or don't try to write to it.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

