---
title: RecentFiles.Maximum Property (Excel)
keywords: vbaxl10.chm172073
f1_keywords:
- vbaxl10.chm172073
ms.prod: excel
api_name:
- Excel.RecentFiles.Maximum
ms.assetid: 24bb3472-8b75-5457-467a-266ed8e5f979
ms.date: 06/08/2017
---


# RecentFiles.Maximum Property (Excel)

Returns or sets the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50. Read/write  **Long** .


## Syntax

 _expression_ . **Maximum**

 _expression_ A variable that represents a **RecentFiles** object.


## Example

This example sets the maximum number of files in the list of recently used files to 6.


```vb
Application.RecentFiles.Maximum = 6
```


## See also


#### Concepts


[RecentFiles Object](recentfiles-object-excel.md)

