---
title: Application.DisplayRecentFiles Property (Word)
keywords: vbawd10.chm158335032
f1_keywords:
- vbawd10.chm158335032
ms.prod: word
api_name:
- Word.Application.DisplayRecentFiles
ms.assetid: d8c96e18-7bbc-baa0-66ae-af91ee631a26
ms.date: 06/08/2017
---


# Application.DisplayRecentFiles Property (Word)

 **True** if the names of recently used files are displayed on the **File** menu. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayRecentFiles**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example sets Word to display a maximum of six file names on the  **File** menu.


```vb
Application.DisplayRecentFiles = True 
RecentFiles.Maximum = 6
```

This example removes the list of recently used files from the  **File** menu.




```vb
Application.DisplayRecentFiles = False
```


## See also


#### Concepts


[Application Object](application-object-word.md)

