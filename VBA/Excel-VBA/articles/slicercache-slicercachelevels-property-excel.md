---
title: SlicerCache.SlicerCacheLevels Property (Excel)
keywords: vbaxl10.chm897079
f1_keywords:
- vbaxl10.chm897079
ms.prod: excel
api_name:
- Excel.SlicerCache.SlicerCacheLevels
ms.assetid: 0fa9bd67-2276-196d-15e6-2570d8c9770a
ms.date: 06/08/2017
---


# SlicerCache.SlicerCacheLevels Property (Excel)

Returns the collection of  **[SlicerCacheLevel](slicercachelevel-object-excel.md)** objects that represent the levels of an OLAP hierarchy on which the specified slicer cache is based. Read-only


## Syntax

 _expression_ . **SlicerCacheLevels**

 _expression_ A variable that represents a **[SlicerCache](slicercache-object-excel.md)** object.


## Remarks

The  **SlicerCacheLevels** property applies only to slicers that filter OLAP data sources ( **SlicerCache** . **[OLAP](slicercache-olap-property-excel.md)** = **True** ). Attempting to access this property from non-OLAP slicers will generate a run-time error.


## See also


#### Concepts


[SlicerCache Object](slicercache-object-excel.md)

