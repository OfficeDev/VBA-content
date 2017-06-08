---
title: TickLabels Object (Excel)
keywords: vbaxl10.chm615072
f1_keywords:
- vbaxl10.chm615072
ms.prod: excel
api_name:
- Excel.TickLabels
ms.assetid: fcb02bc5-fcdc-db32-168b-2d40e5552991
ms.date: 06/08/2017
---


# TickLabels Object (Excel)

Represents the tick-mark labels associated with tick marks on a chart axis.


## Remarks

This object isn't a collection. There's no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit.

Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the  **[TickLabelSpacing](axis-ticklabelspacing-property-excel.md)** property for the category axis.

Tick-mark label text for the value axis is calculated based on the  **[MajorUnit](axis-majorunit-property-excel.md)**, **[MinimumScale](axis-minimumscale-property-excel.md)**, and **[MaximumScale](axis-maximumscale-property-excel.md)** properties of the value axis. To change the tick-mark label text for the value axis, you must change thte values of these properties.


## Example

Use the  **[TickLabels](axis-ticklabels-property-excel.md)** property to return the **TickLabels** object. The following example sets the number format for the tick-mark labels on the value axis in embedded chart one on Sheet1.


```
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Axes(xlValue).TickLabels.NumberFormat = "0.00"
```


## Methods



|**Name**|
|:-----|
|[Delete](ticklabels-delete-method-excel.md)|
|[Select](ticklabels-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](ticklabels-alignment-property-excel.md)|
|[Application](ticklabels-application-property-excel.md)|
|[Creator](ticklabels-creator-property-excel.md)|
|[Depth](ticklabels-depth-property-excel.md)|
|[Font](ticklabels-font-property-excel.md)|
|[Format](ticklabels-format-property-excel.md)|
|[MultiLevel](ticklabels-multilevel-property-excel.md)|
|[Name](ticklabels-name-property-excel.md)|
|[NumberFormat](ticklabels-numberformat-property-excel.md)|
|[NumberFormatLinked](ticklabels-numberformatlinked-property-excel.md)|
|[NumberFormatLocal](ticklabels-numberformatlocal-property-excel.md)|
|[Offset](ticklabels-offset-property-excel.md)|
|[Orientation](ticklabels-orientation-property-excel.md)|
|[Parent](ticklabels-parent-property-excel.md)|
|[ReadingOrder](ticklabels-readingorder-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
