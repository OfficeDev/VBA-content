---
title: ChartGroups Object (Excel)
keywords: vbaxl10.chm569072
f1_keywords:
- vbaxl10.chm569072
ms.prod: excel
api_name:
- Excel.ChartGroups
ms.assetid: 991147bc-bbb5-9f7d-a7c9-55854aa50325
ms.date: 06/08/2017
---


# ChartGroups Object (Excel)

Represents one or more series plotted in a chart with the same format.


## Remarks

 A **ChartGroups** collection is a collection of all the **[ChartGroup](chartgroup-object-excel.md)** objects in the specified chart. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.

Use the  **ChartGroups** method to return the **ChartGroups** collection. The following example displays the number of chart groups on embedded chart 1 on worksheet 1.




```vb
MsgBox Worksheets(1).ChartObjects(1).Chart.ChartGroups.Count
```

Use  **ChartGroups** ( _index_ ), where _index_ is the chart-group index number, to return a single **ChartGroup** object. The following example adds drop lines to chart group 1 on chart sheet 1.




```vb
Charts(1).ChartGroups(1).HasDropLines = True
```

If the chart has been activated, you can use  **ActiveChart** :




```vb
Charts(1).Activate 
ActiveChart.ChartGroups(1).HasDropLines = True
```

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart group shortcut methods to return a particular chart group. The  **PieGroups** method returns the collection of pie chart groups in a chart, the **LineGroups** method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single **ChartGroup** object, or without an index number to return a **ChartGroups** collection.


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

