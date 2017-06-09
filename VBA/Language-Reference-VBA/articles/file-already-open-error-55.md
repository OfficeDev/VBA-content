---
title: File already open (Error 55)
keywords: vblr6.chm1011167
f1_keywords:
- vblr6.chm1011167
ms.prod: office
ms.assetid: cd86a735-910f-5922-3a53-6b9963bb71ae
ms.date: 06/08/2017
---


# File already open (Error 55)

Sometimes a file must be closed before another  **Open** or other operation can occur. This error has the following causes and solutions:



- A sequential-output mode  **Open** statement was executed for a file that is already open. You must close a file opened for one type of sequential access before opening it for another. For example, you must close a file opened for **Input** before opening it for **Output**.
    
- A statement, for example,  **Kill**, **SetAttr**, or **Name**, refers to an open file. Close the file before executing the statement.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

