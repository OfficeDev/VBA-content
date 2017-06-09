---
title: RecentFiles.Item Property (Excel)
keywords: vbaxl10.chm172075
f1_keywords:
- vbaxl10.chm172075
ms.prod: excel
api_name:
- Excel.RecentFiles.Item
ms.assetid: f153bdeb-6c13-2ea8-506a-2b762b211c67
ms.date: 06/08/2017
---


# RecentFiles.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **RecentFiles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

This example opens file two in the list of recently used files.


```vb
Application.RecentFiles.Item(2).Open
```


## See also


#### Concepts


[RecentFiles Object](recentfiles-object-excel.md)

