---
title: PlotArea Object (PowerPoint)
keywords: vbapp10.chm713000
f1_keywords:
- vbapp10.chm713000
ms.prod: powerpoint
api_name:
- PowerPoint.PlotArea
ms.assetid: c1b991b8-8be2-5342-8b5c-814a2e99fec2
ms.date: 06/08/2017
---


# PlotArea Object (PowerPoint)

Represents the plot area of a chart.


## Remarks

 This is the area where your chart data is plotted. The plot area on a 2-D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3-D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart.

The plot area is surrounded by the chart area. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend. For information about formatting the chart area, see the  **[ChartArea](chartarea-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[PlotArea](chart-plotarea-property-powerpoint.md)** property to return a **PlotArea** object. The following example places a dashed border around the chart area of the first chart in the active document, and then places a dotted border around the plot area.




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


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

