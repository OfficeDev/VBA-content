---
title: TimelineState.EndDate Property (Excel)
keywords: vbaxl10.chm950074
f1_keywords:
- vbaxl10.chm950074
ms.prod: excel
ms.assetid: 1d33ce70-32ed-a439-eb34-7305fd9557f2
ms.date: 06/08/2017
---


# TimelineState.EndDate Property (Excel)

Returns the end of the filtering date range (equals to [TimelineState.StartDate Property (Excel)](timelinestate-startdate-property-excel.md) if range is a single day). **Variant** Read-only


## Syntax

 _expression_ . **EndDate**

 _expression_ A variable that represents a[TimelineState](timelinestate-object-excel.md) object.


## Remarks

This property will return an error for either of the following conditions:


- [TimelineState.SingleRangeFilterState Property (Excel)](timelinestate-singlerangefilterstate-property-excel.md) == **False**
    
- [SlicerCache.FilterCleared Property (Excel)](slicercache-filtercleared-property-excel.md) == **True**
    

## Property value

 **VARIANT**


## See also


#### Other resources



[TimelineState Object](timelinestate-object-excel.md)

