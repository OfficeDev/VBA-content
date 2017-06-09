---
title: PivotCaches.Item Method (Excel)
keywords: vbaxl10.chm229074
f1_keywords:
- vbaxl10.chm229074
ms.prod: excel
api_name:
- Excel.PivotCaches.Item
ms.assetid: 80a830fb-a1bf-f1dd-962c-339d99b6f80d
ms.date: 06/08/2017
---


# PivotCaches.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **PivotCaches** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[PivotCache](pivotcache-object-excel.md)** object contained by the collection.


## Example

This example refreshes cache one.


```vb
ActiveWorkbook.PivotCaches.Item(1).Refresh
```


## See also


#### Concepts


[PivotCaches Object](pivotcaches-object-excel.md)

