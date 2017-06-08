---
title: Invalid Access mode
keywords: vblr6.chm1040061
f1_keywords:
- vblr6.chm1040061
ms.prod: office
ms.assetid: a9bb907d-a3e7-993b-9964-d8e6dc163acc
ms.date: 06/08/2017
---


# Invalid Access mode

In your  **Open** statement, you specified a type of access that was invalid for the specified file type. This error has the following causes and solutions:



- You attempted to open a file for  **Input**, but specified an illegal access mode. You can omit the access mode specification when opening a file for input, but if you specify it, the access mode must be **Read**. Both **Write** and **Read Write** are invalid access modes on a file opened for **Input**.
    
- You attempted to open a file for  **Append**, but specified an invalid access mode. You can omit the access mode specification when opening a file for append, but if you specify it, the access mode must be **Write**. Both **Read** and **Read Write** are invalid access modes on a file opened for **Append**.
    
- You attempted to open a file for  **Output**, but specified an invalid access mode. You can omit the access mode specification when opening a file for output, but if you specify it, the access mode must be **Write**. Both **Read** and **Read Write** are invalid access modes on a file opened for **Output**.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

