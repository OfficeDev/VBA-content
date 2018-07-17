---
title: PivotFilters Object (Excel)
keywords: vbaxl10.chm771072
f1_keywords:
- vbaxl10.chm771072
ms.prod: excel
api_name:
- Excel.PivotFilters
ms.assetid: fc647acb-bd6a-8544-6411-1f5e49807e53
ms.date: 06/08/2017
---


# PivotFilters Object (Excel)

The  **PivotFilters** object is a collection of **PivotFilter** objects.


## Remarks

The  **PivotFilters** collection contains properties and methods to add new filters, count the number of existing filters in the collection, and reference specific **PivotFilter** objects.


## Example

In the following example, a new PivotFilter is added to the PivotField at the currently active cell.


```
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek
```


## Methods



|**Name**|
|:-----|
|[Add2](pivotfilters-add-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](pivotfilters-application-property-excel.md)|
|[Count](pivotfilters-count-property-excel.md)|
|[Creator](pivotfilters-creator-property-excel.md)|
|[Item](pivotfilters-item-property-excel.md)|
|[Parent](pivotfilters-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
