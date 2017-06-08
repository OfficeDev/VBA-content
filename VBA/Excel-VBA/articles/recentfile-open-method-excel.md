---
title: RecentFile.Open Method (Excel)
keywords: vbaxl10.chm170076
f1_keywords:
- vbaxl10.chm170076
ms.prod: excel
api_name:
- Excel.RecentFile.Open
ms.assetid: 0db24662-fe68-aa65-1875-0d58f1e37e39
ms.date: 06/08/2017
---


# RecentFile.Open Method (Excel)

Opens a recent workbook.


## Syntax

 _expression_ . **Open**

 _expression_ A variable that represents a **RecentFile** object.


### Return Value

A  **[Workbook](workbook-object-excel.md)** object that represents the opened workbook.


## Example

This example opens the second workbook in the recently used list.


```vb
Application.RecentFiles(2).Open
```


## See also


#### Concepts


[RecentFile Object](recentfile-object-excel.md)

