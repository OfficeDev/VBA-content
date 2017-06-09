---
title: SlicerCache Object (Excel)
keywords: vbaxl10.chm896072
f1_keywords:
- vbaxl10.chm896072
ms.prod: excel
api_name:
- Excel.SlicerCache
ms.assetid: 6e6533e3-0503-a1d3-9ecd-f7997233565f
ms.date: 06/08/2017
---


# SlicerCache Object (Excel)

Represents the current filter state for a slicer and information about which  **[PivotCache](pivotcache-object-excel.md)** or **[WorkbookConnection](workbookconnection-object-excel.md)** the slicer is connected to.


## Remarks

Use the  **[SlicerCaches](workbook-slicercaches-property-excel.md)** property of the **[Workbook](workbook-object-excel.md)** object to access the collection of **SlicerCache** objects in a workbook.

Each slicer has a base  **SlicerCache** object which represents the items displayed in the slicer and the current user interface state of the tiles displayed with their corresponding item captions. Each slicer control that the user sees in Excel is represented by a **[Slicer](slicer-object-excel.md)** object that has a **SlicerCache** object associated with it.


## Example

The following code example creates a  **SlicerCache** object based on the Customer Geography OLAP hierarchy from the connection to the AdventureWorks database, and then creates a slicer on the Country level of that hierarchy in Sheet2 of the workbook.


```
With ActiveWorkbook 
 .SlicerCaches.Add("AdventureWorks", _ 
 "[Customer].[Customer Geography]").Slicers.Add SlicerDestination:="Sheet2", _ 
 Level:="[Customer].[Customer Geography].[Country]", Caption:="Country" 
End With 

```


## Methods



|**Name**|
|:-----|
|[ClearAllFilters](slicercache-clearallfilters-method-excel.md)|
|[ClearDateFilter](slicercache-cleardatefilter-method-excel.md)|
|[ClearManualFilter](slicercache-clearmanualfilter-method-excel.md)|
|[Delete](slicercache-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](slicercache-application-property-excel.md)|
|[Creator](slicercache-creator-property-excel.md)|
|[CrossFilterType](slicercache-crossfiltertype-property-excel.md)|
|[FilterCleared](slicercache-filtercleared-property-excel.md)|
|[Index](slicercache-index-property-excel.md)|
|[List](slicercache-list-property-excel.md)|
|[ListObject](slicercache-listobject-property-excel.md)|
|[Name](slicercache-name-property-excel.md)|
|[OLAP](slicercache-olap-property-excel.md)|
|[Parent](slicercache-parent-property-excel.md)|
|[PivotTables](slicercache-pivottables-property-excel.md)|
|[RequireManualUpdate](slicercache-requiremanualupdate-property-excel.md)|
|[ShowAllItems](slicercache-showallitems-property-excel.md)|
|[SlicerCacheLevels](slicercache-slicercachelevels-property-excel.md)|
|[SlicerCacheType](slicercache-slicercachetype-property-excel.md)|
|[SlicerItems](slicercache-sliceritems-property-excel.md)|
|[Slicers](slicercache-slicers-property-excel.md)|
|[SortItems](slicercache-sortitems-property-excel.md)|
|[SortUsingCustomLists](slicercache-sortusingcustomlists-property-excel.md)|
|[SourceName](slicercache-sourcename-property-excel.md)|
|[SourceType](slicercache-sourcetype-property-excel.md)|
|[TimelineState](slicercache-timelinestate-property-excel.md)|
|[VisibleSlicerItems](slicercache-visiblesliceritems-property-excel.md)|
|[VisibleSlicerItemsList](slicercache-visiblesliceritemslist-property-excel.md)|
|[WorkbookConnection](slicercache-workbookconnection-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
