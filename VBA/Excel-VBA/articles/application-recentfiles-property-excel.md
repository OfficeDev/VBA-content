---
title: Application.RecentFiles Property (Excel)
keywords: vbaxl10.chm133170
f1_keywords:
- vbaxl10.chm133170
ms.prod: excel
api_name:
- Excel.Application.RecentFiles
ms.assetid: a64784af-4162-90fc-b955-963a1b1e747f
ms.date: 06/08/2017
---


# Application.RecentFiles Property (Excel)

Returns a  **[RecentFiles](recentfiles-object-excel.md)** collection that represents the list of recently used files.


## Syntax

 _expression_ . **RecentFiles**

 _expression_ A variable that represents an **Application** object.


## Example

This example sets the maximum number of files in the list of recently used files to 6.


```vb
Application.RecentFiles.Maximum = 6
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

