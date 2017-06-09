---
title: TimelineState.StartDate Property (Excel)
keywords: vbaxl10.chm950073
f1_keywords:
- vbaxl10.chm950073
ms.prod: excel
ms.assetid: 3de8df53-1a36-428e-50dd-c7f45aa73b25
ms.date: 06/08/2017
---


# TimelineState.StartDate Property (Excel)

Returns the start of the filtering date range.  **Variant** Read-only


## Syntax

 _expression_ . **StartDate**

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

