---
title: SlicerCache.SlicerItems Property (Excel)
keywords: vbaxl10.chm897083
f1_keywords:
- vbaxl10.chm897083
ms.prod: excel
api_name:
- Excel.SlicerCache.SlicerItems
ms.assetid: d552a519-3d9f-74b8-4cbe-3b5c935a14d9
ms.date: 06/08/2017
---


# SlicerCache.SlicerItems Property (Excel)

Returns a  **[SlicerItems](sliceritems-object-excel.md)** collection that contains the collection of all items in the slicer cache. Read-only


## Syntax

 _expression_ . **SlicerItems**

 _expression_ A variable that represents a **[SlicerCache](slicercache-object-excel.md)** object.


### Return Value

 **SlicerItems**


## Remarks

The  **SlicerItems** property of the **SlicerCache** object is only applicable for slicers that are based on PivotTables based on workbook ranges or lists ( **SlicerCache** . **SourceType** = **xlDatabase** ), or for slicers that are based on PivotTables based on relational data sources ( **SlicerCache** . **SourceType** = **xlExternal** and **SlicerCache** . **[OLAP](slicercache-olap-property-excel.md)** = **False** ). Attempting to access the **SlicerItems** property for slicers that are connected to an external OLAP data source ( **SlicerCache** . **OLAP** = **True** ) generates a run-time error. For OLAP data sources, use the **[SlicerItems](slicercachelevel-sliceritems-property-excel.md)** property of the **[SlicerCacheLevel](slicercachelevel-object-excel.md)** object instead.


## See also


#### Concepts


[SlicerCache Object](slicercache-object-excel.md)

