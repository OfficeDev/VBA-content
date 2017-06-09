---
title: Series.PlotOrder Property (PowerPoint)
keywords: vbapp10.chm65764
f1_keywords:
- vbapp10.chm65764
ms.prod: powerpoint
api_name:
- PowerPoint.Series.PlotOrder
ms.assetid: 196c0b37-a9fe-ec01-ca0a-786c70e8a63c
ms.date: 06/08/2017
---


# Series.PlotOrder Property (PowerPoint)

Returns or sets the plot order for the selected series within the chart group. Read/write  **Long**.


## Syntax

 _expression_. **PlotOrder**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Remarks

You can set plot order only within a chart group (you cannot set the plot order for the entire chart if you have more than one chart type). A chart group is a collection of series with the same chart type.

Changing the plot order of one series will cause the plot orders of the other series in the chart group to be adjusted, as necessary.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example makes series two for the first chart in the active document appear third in the plot order. You should run the example on a 2-D column chart that contains three or more series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).SeriesCollection(2).PlotOrder = 3

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

