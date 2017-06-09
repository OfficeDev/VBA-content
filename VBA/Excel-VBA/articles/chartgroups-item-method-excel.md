---
title: ChartGroups.Item Method (Excel)
keywords: vbaxl10.chm570074
f1_keywords:
- vbaxl10.chm570074
ms.prod: excel
api_name:
- Excel.ChartGroups.Item
ms.assetid: 29ca6f13-96b7-bd43-9562-480c467ef7db
ms.date: 06/08/2017
---


# ChartGroups.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **ChartGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

### Return Value

A  **[ChartGroup](chartgroup-object-excel.md)** object contained by the collection.


## Example

This example adds drop lines to chart group one on chart sheet one.


```vb
Charts(1).ChartGroups.Item(1).HasDropLines = True
```


## See also


#### Concepts


[ChartGroups Object](chartgroups-object-excel.md)

