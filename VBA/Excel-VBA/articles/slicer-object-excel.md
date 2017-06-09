---
title: Slicer Object (Excel)
keywords: vbaxl10.chm904072
f1_keywords:
- vbaxl10.chm904072
ms.prod: excel
api_name:
- Excel.Slicer
ms.assetid: 577be0f6-4eda-0093-8899-097f3c900383
ms.date: 06/08/2017
---


# Slicer Object (Excel)

Represents a slicer in a workbook.


## Remarks

Each  **Slicer** object represents a slicer in a workbook. Slicers are used to filter data in PivotTable reports or OLAP data sources.

Use the  **[Add](slicers-add-method-excel.md)** method to add a **Slicer** object to the **[Slicers](slicers-object-excel.md)** collection. To access the **SlicerItem** object that represents the currently selected button in a slicer, use the **[ActiveItem](slicer-activeitem-property-excel.md)** property of the **Slicer** object.


## Example

The following code example changes the caption for the first slicer in the first slicer cache to "My Slicer".


```
ActiveWorkbook.SlicerCaches(1).Slicers(1).Caption = "My Slicer"
```

The following code example sets the width of the first slicer in the first slicer cache to equal 200 points.




```
ActiveWorkbook.SlicerCaches(1).Slicers(1).Width = 200
```


## Methods



|**Name**|
|:-----|
|[Copy](slicer-copy-method-excel.md)|
|[Cut](slicer-cut-method-excel.md)|
|[Delete](slicer-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[ActiveItem](slicer-activeitem-property-excel.md)|
|[Application](slicer-application-property-excel.md)|
|[Caption](slicer-caption-property-excel.md)|
|[ColumnWidth](slicer-columnwidth-property-excel.md)|
|[Creator](slicer-creator-property-excel.md)|
|[DisableMoveResizeUI](slicer-disablemoveresizeui-property-excel.md)|
|[DisplayHeader](slicer-displayheader-property-excel.md)|
|[Height](slicer-height-property-excel.md)|
|[Left](slicer-left-property-excel.md)|
|[Locked](slicer-locked-property-excel.md)|
|[Name](slicer-name-property-excel.md)|
|[NumberOfColumns](slicer-numberofcolumns-property-excel.md)|
|[Parent](slicer-parent-property-excel.md)|
|[RowHeight](slicer-rowheight-property-excel.md)|
|[Shape](slicer-shape-property-excel.md)|
|[SlicerCache](slicer-slicercache-property-excel.md)|
|[SlicerCacheLevel](slicer-slicercachelevel-property-excel.md)|
|[SlicerCacheType](slicer-slicercachetype-property-excel.md)|
|[Style](slicer-style-property-excel.md)|
|[TimelineViewState](slicer-timelineviewstate-property-excel.md)|
|[Top](slicer-top-property-excel.md)|
|[Width](slicer-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
