---
title: ChartGroups Object (PowerPoint)
keywords: vbapp10.chm693000
f1_keywords:
- vbapp10.chm693000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroups
ms.assetid: 2db874db-91af-0b1e-7496-92a8443caade
ms.date: 06/08/2017
---


# ChartGroups Object (PowerPoint)

Represents one or more series plotted in a chart with the same format.


## Remarks

 A **ChartGroups** collection is a collection of all the **[ChartGroup](chartgroup-object-powerpoint.md)** objects in the specified chart. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

 The following example displays the number of chart groups on the first chart of the active document. Use the **[ChartGroups](chart-chartgroups-method-powerpoint.md)** method to return the **ChartGroups** collection.




```vb
MsgBox ActiveDocument.InlineShapes(1).Chart._

    ChartGroups.Count
```

The following example adds drop lines to chart group 1 on chart sheet 1. Use  **ChartGroups** ( _index_ ), where _Index_ is the chart group index number, to return a single **ChartGroup** object.




```vb
ActiveDocument.InlineShapes(1).Chart._

    ChartGroups(1).HasDropLines = True
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

