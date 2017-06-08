---
title: ChartGroups Object (Word)
ms.prod: word
api_name:
- Word.ChartGroups
ms.assetid: 37136fbd-8740-c817-9666-993bc5d4c847
ms.date: 06/08/2017
---


# ChartGroups Object (Word)

Represents one or more series plotted in a chart with the same format.


## Remarks

 A **ChartGroups** collection is a collection of all the **[ChartGroup](chartgroup-object-word.md)** objects in the specified chart. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.

 The following example displays the number of chart groups on the first chart of the active document. Use the **[ChartGroups](chart-chartgroups-property-word.md)** property to return the **ChartGroups** collection.




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


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

