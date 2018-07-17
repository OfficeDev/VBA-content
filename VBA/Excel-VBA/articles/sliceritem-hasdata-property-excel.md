---
title: SlicerItem.HasData Property (Excel)
keywords: vbaxl10.chm907080
f1_keywords:
- vbaxl10.chm907080
ms.prod: excel
api_name:
- Excel.SlicerItem.HasData
ms.assetid: 17ce0cdc-ec30-638a-e869-4640ee0ef5a3
ms.date: 06/08/2017
---


# SlicerItem.HasData Property (Excel)

Returns whether the slicer item contains data that matches the current manual filter state. Read-only.


## Syntax

 _expression_ . **HasData**

 _expression_ A variable that represents a **[SlicerItem](sliceritem-object-excel.md)** object.


### Return Value

Boolean


## Remarks

 **True** if the slicer item contains data that matches the current manual filter state; otherwise **False** .

This property is only supported when cross filtering is turned on, and will generate a run-time error otherwise. To determine if cross filtering is turned on for a slicer associated with a PivotTable report, check the value of the  **[CrossFilterType](slicercache-crossfiltertype-property-excel.md)** property of the specified **SlicerItem** object's parent **[SlicerCache](slicercache-object-excel.md)** object. To determine if cross filtering is turned on for a slicer associated with an OLAP data source, check the value of the **[CrossFilterType](slicercachelevel-crossfiltertype-property-excel.md)** property of the **[SlicerCacheLevel](slicercachelevel-object-excel.md)** object that corresponds to the OLAP hierarchy being filtered.


## See also


#### Concepts


[SlicerItem Object](sliceritem-object-excel.md)

