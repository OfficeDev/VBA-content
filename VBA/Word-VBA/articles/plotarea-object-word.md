---
title: PlotArea Object (Word)
keywords: vbawd10.chm816
f1_keywords:
- vbawd10.chm816
ms.prod: word
api_name:
- Word.PlotArea
ms.assetid: 72d30767-7cfc-3063-0b49-f9fbc129a52c
ms.date: 06/08/2017
---


# PlotArea Object (Word)

Represents the plot area of a chart.


## Remarks

 This is the area where your chart data is plotted. The plot area on a 2-D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3-D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart.

The plot area is surrounded by the chart area. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend. For information about formatting the chart area, see the  **[ChartArea](chartarea-object-word.md)** object.


## Example

Use the  **[PlotArea](chart-plotarea-property-word.md)** property to return a **PlotArea** object. The following example places a dashed border around the chart area of the first chart in the active document, and then places a dotted border around the plot area.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
 End With 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


