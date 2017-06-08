---
title: SlicerCacheLevel.CrossFilterType Property (Excel)
keywords: vbaxl10.chm901077
f1_keywords:
- vbaxl10.chm901077
ms.prod: excel
api_name:
- Excel.SlicerCacheLevel.CrossFilterType
ms.assetid: 2e91a528-a253-e731-6f11-a33f9ee6d9e2
ms.date: 06/08/2017
---


# SlicerCacheLevel.CrossFilterType Property (Excel)

Returns or sets whether a slicer is participating in cross filtering with other slicers that share the same slicer cache, and how cross filtering is displayed. Read/write


## Syntax

 _expression_ . **CrossFilterType**

 _expression_ A variable that represents a **[SlicerCacheLevel](slicercachelevel-object-excel.md)** object.


### Return Value

 **[XlSlicerCrossFilterType](xlslicercrossfiltertype-enumeration-excel.md)**


## Remarks

If more than one slicer is associated with the same PivotTable, by default, if the item or items you filter by in one slicer have no corresponding data in another slicer, those items will be dimmed. For example, if you have Country slicer and a State slicer, and you click a country in the Country slicer, all states in the State slicer that are not in that country will be dimmed. This feature is referred to as "cross filtering."

The user interface settings that correspond to the setting of the  **CrossFilterType** property are the **Visually indicate items with no data** and **Show items with no data last** check boxes in the **Slicer Settings** dialog box. Setting the **CrossFilterType** property to **xlSlicerCrossFilterShowItemsWithDataAtTop** corresponds to selecting both the **Visually indicate items with no data** and **Show items with no data last** check boxes. Setting the **CrossFilterType** property to **xlSlicerCrossFilterShowItemsWithNoData** corresponds to selecting only the **Visually indicate items with no data** check box. Clearing both check boxes corresponds to setting the **CrossFilterType** property to **xlSlicerNoCrossFilter** .

Slicers filtering non-OLAP data sources ( **SlicerCache** . **OLAP** = **False** ) are not supported by the **CrossFilterType** property of the **SlicerCacheLevel** object. For slicers filtering PivotTables based on ranges, lists, and non-OLAP relational data sources, use the **[CrossFilterType](slicercache-crossfiltertype-property-excel.md)** property of the **[SlicerCache](slicercache-object-excel.md)** object, instead.


## See also


#### Concepts


[SlicerCacheLevel Object](slicercachelevel-object-excel.md)

