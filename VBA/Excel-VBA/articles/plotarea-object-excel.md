---
title: PlotArea Object (Excel)
keywords: vbaxl10.chm617072
f1_keywords:
- vbaxl10.chm617072
ms.prod: excel
api_name:
- Excel.PlotArea
ms.assetid: 85c42124-268c-8b0e-ba5d-c2f6fbf53e79
ms.date: 06/08/2017
---


# PlotArea Object (Excel)

Represents the plot area of a chart.


## Remarks

 This is the area where your chart data is plotted. The plot area on a 2-D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3-D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart.

The plot area is surrounded by the chart area. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend. For information about formatting the chart area, see the  **[ChartArea](chartarea-object-excel.md)** object.


## Example

Use the  **PlotArea** property to return a **PlotArea** object. The following example activates the chart sheet named "Chart1," places a dashed border around the chart area of the active chart, and places a dotted border around the plot area.


```
Charts("Chart1").Activate 
With ActiveChart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
End With
```


## Methods



|**Name**|
|:-----|
|[ClearFormats](plotarea-clearformats-method-excel.md)|
|[Select](plotarea-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](plotarea-application-property-excel.md)|
|[Creator](plotarea-creator-property-excel.md)|
|[Format](plotarea-format-property-excel.md)|
|[Height](plotarea-height-property-excel.md)|
|[InsideHeight](plotarea-insideheight-property-excel.md)|
|[InsideLeft](plotarea-insideleft-property-excel.md)|
|[InsideTop](plotarea-insidetop-property-excel.md)|
|[InsideWidth](plotarea-insidewidth-property-excel.md)|
|[Left](plotarea-left-property-excel.md)|
|[Name](plotarea-name-property-excel.md)|
|[Parent](plotarea-parent-property-excel.md)|
|[Position](plotarea-position-property-excel.md)|
|[Top](plotarea-top-property-excel.md)|
|[Width](plotarea-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
