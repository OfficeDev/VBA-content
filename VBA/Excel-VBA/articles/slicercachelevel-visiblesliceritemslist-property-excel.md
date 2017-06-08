---
title: SlicerCacheLevel.VisibleSlicerItemsList Property (Excel)
keywords: vbaxl10.chm901079
f1_keywords:
- vbaxl10.chm901079
ms.prod: excel
api_name:
- Excel.SlicerCacheLevel.VisibleSlicerItemsList
ms.assetid: 68c0800b-4130-59f2-d0c0-7cad49b98f0d
ms.date: 06/08/2017
---


# SlicerCacheLevel.VisibleSlicerItemsList Property (Excel)

Returns the list of slicer items that are currently included in the slicer filter. Read-only


## Syntax

 _expression_ . **VisibleSlicerItemsList**

 _expression_ A variable that represents a **[SlicerCacheLevel](slicercachelevel-object-excel.md)** object.


### Return Value

 **Variant**


## Remarks

The list of slicer items are returned as MDX unique name strings. If this list is empty, the slicer is not filtering the data source and all slicer tiles are displayed as selected.


## See also


#### Concepts


[SlicerCacheLevel Object](slicercachelevel-object-excel.md)

